Grok Speech Recognition using the API via Powershell

To use your API Key, create a system environment variable named GrokAPIKey with your key information<br>
Example of how to add your API key with an Admin CMD prompt:

SETX GrokAPIKey Key01010101010Data0101010101010Here

Then close the console window and run Grokory.ps1(the env var key wont reflect in the console it is set in, only new/future consoles), if no API key is found a canned response will be given when Grok speaks.

Hold the left Ctrl key to speak to Grok, release the key and he will speak back to you. ;)<br>
Initial version is console based, eventually this will be GUI based. This is an extremely rough first draft.

Made with love, by Grok 3
