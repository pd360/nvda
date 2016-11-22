/*
Header for C++ code to provide access to Windows OneCore voices.
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

#pragma once
#define export __declspec(dllexport) 

typedef int (*ocSpeech_Callback)(byte* data, int length, const char16* markers);
extern "C" {
void export __stdcall ocSpeech_initialize();
void export __stdcall ocSpeech_setCallback(ocSpeech_Callback fn);
export int __stdcall ocSpeech_speak(char16 *s);
export const wchar_t * __stdcall ocSpeech_getVoices(void);
export void __stdcall ocSpeech_setVoice(int i);
export void __stdcall ocSpeech_setProperty(char16 *name, long val);
export const char16 * __stdcall ocSpeech_getCurrentVoiceLanguage();
}
