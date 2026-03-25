HOW TO BUILD THE INSTALLER
==========================

PREREQUISITES
  - Your meeting recorder app must be working
  - You must be in C:\meeting_recorder with the venv activated


STEPS
-----

1. Copy this entire "installer" folder into:
   C:\meeting_recorder\installer\

2. Double-click build.bat
   (Right-click → Run as Administrator for best results)

3. Wait 3-5 minutes while it:
   - Reads all your app source files
   - Embeds them into the installer
   - Compiles everything into one .exe

4. Your installer is at:
   C:\meeting_recorder\dist\MeetingRecorderSetup.exe

5. Share that file with your team!


WHAT THE INSTALLER DOES FOR YOUR TEAM
--------------------------------------

Screen 1 - Welcome
  Overview of what will be installed

Screen 2 - System Check
  Automatically detects:
  ✓ Administrator rights
  ✓ Disk space (requires 5GB free)
  ✓ Nvidia GPU (auto-selects GPU or CPU PyTorch)
  ✓ Internet connection

Screen 3 - Install Location
  User chooses where to install (default: home/meeting_recorder)

Screen 4 - Python 3.11
  Checks for Python 3.11 specifically
  If not found: shows download link with step-by-step instructions
  User installs Python then clicks "Check Again"

Screen 5 - Anthropic API Key
  Step-by-step guide to create account and get API key
  Link opens Anthropic Console in browser
  Key is validated BEFORE install starts (no wasted time)

Screen 6 - HuggingFace Token
  Step-by-step guide to create free account and get token
  Link opens HuggingFace in browser
  Token is validated BEFORE install starts

Screen 7 - Model Terms
  Step-by-step instructions to accept terms for both models:
  - pyannote/speaker-diarization-3.1
  - pyannote/segmentation-3.0
  Direct links open the pages in browser
  Installer VERIFIES access before proceeding

Screen 8 - Installing
  Progress bar showing each step
  Full log so they can see what's happening
  Clear error messages if anything fails

Screen 9 - Done
  Checklist of what was installed
  Launch button to open the app immediately


EDGE CASES HANDLED
------------------
✓ GPU detection (auto CUDA vs CPU PyTorch)
✓ PyAudio fallback via pipwin if C++ Build Tools missing
✓ Corporate firewall detection with IT instructions
✓ numpy version pinning for compatibility
✓ All PyTorch 2.6 / numpy 2.0 patches applied to main.py
✓ HuggingFace token validation before install
✓ Anthropic key validation before install
✓ pyannote model term verification before install
✓ Retry button if install fails
✓ README.txt written to install folder
✓ SmartScreen warning explained on Done screen


WINDOWS SMARTSCREEN WARNING
---------------------------
When your team runs MeetingRecorderSetup.exe they may see:
"Windows protected your PC"

This is normal for unsigned .exe files. Tell them:
1. Click "More info"
2. Click "Run anyway"

If your company has strict IT policies, you may need to
get the .exe whitelisted or signed with a code signing cert.
