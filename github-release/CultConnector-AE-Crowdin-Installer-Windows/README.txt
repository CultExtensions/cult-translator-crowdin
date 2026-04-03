Cult Connector (AE ↔ Crowdin)
=============================

Cult Connector is a ScriptUI panel for Adobe After Effects. It connects your compositions to Crowdin so you can send strings with screenshots, translate with full visual context, and import localized compositions—without leaving your timeline for the messy parts.

Why teams use it
----------------
• Stop exporting strings by hand. Send the right text, with the right frame, to the right project—directly from your After Effects panel.

• Give translators real context. Every string can include a screenshot from your composition, so meaning, tone, and layout stay accurate.

• Bring translations back instantly. Import from Crowdin and get ready-to-review compositions per language—no rebuilding, no guesswork.

Feature highlights
------------------
🎬 Work inside After Effects — No switching tools. Open the panel, log in via browser, and connect to your Crowdin project in seconds.

⚙️ Clear Composition & Settings workflow — Control segmentation, target languages, and project connections through simple, dedicated tabs—built for production speed.

🖼️ Strings + screenshots together — Every text layer is sent with visual context when needed—so translators and reviewers understand exactly how it appears on screen.

🎯 Frame-accurate markers (optional) — For complex motion, markers define the exact frame used for screenshots—ensuring precise context in dense timelines.

🌍 Import translated comps — Pull translations back and automatically generate compositions—typically one per language—ready for fast review and delivery.

How it works (overview)
------------------------
1. Install Cult Connector and open it inside After Effects.
2. Connect to Crowdin (Teams or Enterprise).
3. Send compositions—text + screenshots sync automatically.
4. Translate in Crowdin with full visual context.
5. Import and get localized comps instantly in After Effects.

Requirements
------------
• Windows 10 or Windows 11 (64-bit)
• Adobe After Effects installed (for example “Adobe After Effects 2025” or “2026” from the Creative Cloud desktop app)
• A Cult Extensions account (see “Create your account” below) and a Crowdin project you can access

What is in this package?
------------------------
• Install.vbs — in the folder OUTSIDE SupportFiles (the only file there). Double-click this to install.
• SupportFiles — this folder holds the README, the panel (.jsxbin), the installer script, and optional Install.bat.

Install on Windows (recommended)
--------------------------------
1. Quit After Effects completely (check the system tray so no background instance stays open).

2. Open the parent folder (the one that contains both Install.vbs and this SupportFiles folder). Double-click Install.vbs. No command window stays open; only the guided installer appears. Follow the on-screen steps.

3. The installer copies “Cult Connector (AE ↔ Crowdin).jsxbin” into every After Effects “ScriptUI Panels” folder it finds under your user profile and, if possible, under Program Files. If those folders do not exist yet, it creates them. Any previous Cult Connector .jsxbin in those folders is removed first.

4. If Windows cannot write under Program Files, right-click Install.vbs and choose “Run as administrator”, or rely on the copy under your user profile only (enough for most setups).

5. Launch After Effects again.

6. Open the panel from the menu: Window → Cult Connector (AE ↔ Crowdin) (or the entry that matches the script name).

If you install a new major version of After Effects later, run Install.vbs again.

Optional: from inside SupportFiles you can double-click Install.bat to launch the same installer (it starts the Install.vbs in the parent folder).

Manual install (optional)
-------------------------
1. Quit After Effects.

2. From this SupportFiles folder, copy “Cult Connector (AE ↔ Crowdin).jsxbin” into the ScriptUI Panels folder for your After Effects version. Create the folder if it does not exist. Typical locations:

   Per user (recommended):
   %APPDATA%\Adobe\After Effects\<version>\Scripts\ScriptUI Panels\
   Example: C:\Users\<YourName>\AppData\Roaming\Adobe\After Effects\26.0\Scripts\ScriptUI Panels\

   All users / application folder (optional; may require Administrator when pasting):
   C:\Program Files\Adobe\Adobe After Effects <year>\Support Files\Scripts\ScriptUI Panels\

   <version> is a number such as 24.0, 25.0, 26.0 (it does not always match the marketing year in the Start Menu name).

3. Relaunch After Effects and open the panel under the Window menu.

After Effects setup (best practices — do this first)
---------------------------------------------------
These preferences are required for most workflow automation panels that talk to the network or write files (screenshots, temp data, exports).

1. In After Effects, choose Edit → Preferences → Scripting & Expressions (in some builds this appears under Edit → Settings → Scripting & Expressions).

2. Enable “Allow Scripts to Write Files and Access Network”.

3. Click OK, then fully quit and restart After Effects so the change applies everywhere.

Optional but helpful:
• If the app warns before running scripts, you can adjust “Warn User When Executing Scripts” in the same Scripting preferences area—some teams disable the warning for daily production use once they trust the panel.
• Keep After Effects updated to a supported year build; odd scripting bugs are often fixed in patches.

Using Cult Connector (quick guide)
----------------------------------
1. Open Window → Cult Connector (AE ↔ Crowdin) and dock the panel where you like (often next to other pipeline tools).

2. Sign in: Use the panel’s sign-in flow. Your browser opens for authentication with Cult Extensions / Crowdin. Stay signed in in the browser when the panel expects it—OAuth and session handoffs rely on that.

3. Connect Crowdin: Choose Teams or Enterprise when prompted (Enterprise is supported). Link the Crowdin project you want this AE project to talk to.

4. Composition tab: Pick the compositions (and text layers as needed) you want to send. Review segmentation if your workflow needs finer control over how strings are split before they go to Crowdin.

5. Screenshots / context: The panel can send screenshots with strings. For tricky motion, add composition markers on the exact frame you want captured—markers are optional but recommended when “the right frame” is not obvious.

6. Settings tab: Confirm target languages, project connection, and any options your team uses before syncing.

7. Send / sync: Push strings (and screenshot context) to Crowdin. Translators work there with the same visuals you approved.

8. Import: When translations are ready, import from the panel. After Effects typically gets one composition per language (or equivalent layout your build uses) so you can review and deliver fast.

9. Save your AE project after large imports, and keep network/VPN stable during sync—the panel talks to remote services.

Create your account
-------------------
Start your 14-day free trial — Create an account to manage your license, team access, and subscription in one place, then connect After Effects to Crowdin in minutes.

→ https://cultextensions.com

Sign-in and billing are handled through the Cult Extensions dashboard. After Effects connects through the Cult Connector panel only after your account and Crowdin access are set up.

FAQ
---
Does this work with Crowdin Enterprise?
Yes—choose Enterprise when signing in from the panel.

Do I need markers?
No. They’re optional, but helpful for complex compositions where a clear screenshot frame isn’t obvious.

What gets sent to Crowdin?
Text from your selected compositions, paired with screenshots for accurate translation context (when your workflow uses screenshots).

Is there a free trial?
Yes—new accounts include a 14-day free trial with full functionality.

Support
-------
contact@cultextensions.com
