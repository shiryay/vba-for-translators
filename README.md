This repo contains some VBA macros for MS Word that I find useful in my work as a translator.

The SearchModule.bas (dependent on UtilityModule.bas) contains functions to find the selected text on the following websites:
Google, Google Translate, Linguee (English, Russian, French, Spanish), Proz, insur-info.ru, ozdic.com, multitran.com, abkuerzungen.de, acronymfinder.com
If no text is selected, an input box opens. The search text is URL-encoded using the URLEncode function borrowed from https://excelvba.ru/code/URLEncode

The utility module contains functions for removing hidden text in a Word document, pasting copied text with the original formatting, and converting dates from British to American style and vice-versa.
