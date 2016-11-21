#pragma once
#define export __declspec(dllexport) 

typedef int (*callbackType)(byte* data, int length, const char16* markers);
extern "C" {
void export __stdcall initialize();
void export __stdcall set_callback(callbackType fn);
export int __stdcall speak(char16 *s);
export const wchar_t * __stdcall get_voices(void);
export void __stdcall set_voice(int i);
export void __stdcall set_property(char16 *name, long val);
export const char16 * __stdcall get_current_voice_language();
}
