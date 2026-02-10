# 🚀 Chrome Extensions - REMAS by Zulali
### **R**ecover • **E**xport • **M**erge • **A**udit • **S**ort

**REMAS** is the ultimate power-user utility for managing Chrome extension chaos. Whether you are migrating across machines, merging work and personal profiles, or rescuing "lost" extensions from a backup, REMAS automates the entire lifecycle of extension management.

---

## 🛠 What is REMAS? (The Process)
Managing 100+ extensions across different profiles is usually a manual nightmare. **REMAS** turns this into a 10-second automated process:

1.  🔍 **Deep Scan:** The engine dives into your raw Chrome data, bypassing cryptic 32-character folder names (IDs).
2.  📂 **Structure Flattening:** It extracts the active source code from nested version sub-folders and brings them to the surface.
3.  ⚖️ **Version Auditing:** If an extension exists in both sources, REMAS compares the version strings (e.g., `v2.4.1` vs `v2.4.0`) and **only keeps the latest one**.
4.  🏷 **Human-Readable Renaming:** It reads the internal `manifest.json` (and locale files) to rename folders from gibberish to their actual names (e.g., `ghbmn...` becomes `Volume Master`).
5.  🖥 **Dashboard Generation:** It builds a custom **Extensions.html** "Launchpad" for one-click recovery, store-searching and sideloading.

---

## 💎 Why Power Users Need This
If you have 100+ extensions, you know the pain of "Profile Drift." 

*   **⚡️ Profile Merging:** Finally combine your curated list of extensions from your **Work Laptop** and **Home PC** into one master profile.
*   **🛡 Rescue "Deleted" Tools:** If an extension is removed from the Web Store, you can’t get it back—unless you have the raw source. REMAS saves that source code permanently.
*   **🔓 Manifest V2 Continuity:** For users who rely on legacy MV2 tools (like specific ad-blockers or scrapers), REMAS provides the "Load Unpacked" path to keep them running in Chrome 144 and Latest.
*   **📋 Clean Audit:** Stop wondering which profile has the newest version. REMAS identifies the winner for you.

---

## 🚀 Quick Start Guide

### 1. 📂 Preparation
*   **Single Source:** Place your backup folder on your desktop and rename it to **Extensions_1**. Ensure the folder contain the raw ID folders (e.g., `pkedc...`, `ghbko...`) directly.
*   **Dual-Source Merge:** If you want to merge two different profiles (e.g., your **Work PC** and **Home PC**), similarly place the second backup folder and rename it to **Extensions_2**.
*   **Intelligent Fallback:** REMAS features an automated "All-or-Nothing" logic. If **Extensions_1** is not detected on your desktop, the engine assumes you are auditing your current machine and will skip the desktop search entirely. It will instead fallback to your live Chrome system directory: 
    `...\AppData\Local\Google\Chrome\User Data\Default\Extensions`

**Note:** In the generated `Extensions.html` dashboard, **Source 1** will represent your primary desktop backup (if found) or your **Default Chrome Profile** (in Fallback Mode).

> **⚠️ POWER-USER TIP:** If you are auditing your **Current Chrome Extensions** (Live Default Folder), please **Close Chrome completely** before running the engine. This prevents "File-In-Use" errors and ensures the audit can access your live extension manifest data without being blocked by the browser.

### 2. ⚙️ Configuration
Open the **REMAS script** in **Notepad** (or any text editor) and change the `$user` variable to match your Windows user folder name. 
*   **Example:** If your Windows user folder is `C:\Users\JohnDoe`, change the line to: `$user = "JohnDoe"`.
*   Once edited, **Select All (Ctrl+A)** and **Copy (Ctrl+C)** the entire script.

### 3. ⚡ Execute the REMAS Process
Open **PowerShell as Administrator**, paste the entire copied script, and hit **Enter**. The engine will immediately begin auditing your folders and will generate the **Chrome Extensions - REMAS** folder on your desktop.

### 4. 🔓 Unlock Chrome 144 or Latest (The Bypasses for Sideloading)
Google blocks legacy extensions by default. Use these to stay in control:

**Why is this required?** Google has officially deprecated Manifest V2 extensions in Chrome 144+. Without these bypasses, Chrome will automatically disable your recovered legacy extensions and block you from re-enabling them. Implementing either (or both) of these ensures you stay in control of your software:

*   **Option A: Registry Fix (Recommended - Best for "Set & Forget"):**
    This is the most reliable method. It applies a system-wide policy so your extensions work no matter how you launch Chrome.
    ```powershell
    # Forces Chrome to allow Manifest V2 support
    $path = "HKLM:\SOFTWARE\Policies\Google\Chrome"
    if (!(Test-Path $path)) { New-Item -Path $path -Force }
    New-ItemProperty -Path $path -Name "ExtensionManifestV2Availability" -Value 2 -PropertyType DWORD -Force
    ```
*   **Option B: Shortcut Target Fix (Best for non-Admin users):**
    Use this if you cannot edit the Registry. It enables the Legacy feature directly in the browser engine.
    Add this to the end of your Chrome shortcut **Target** field (after the quotes):
    ```text
    --enable-features=AllowLegacyMV2Extensions
    ```
**Note:** For the most stable experience in 2026, implementing **both** is highly recommended to bypass both policy checks and engine-level blocks simultaneously.

#### **Which one is the "Better" option?**
The **Registry Fix** is the superior, "Set and Forget" option. 
* **The Winner (Registry):** It applies a system-wide policy. Once you run the PowerShell command, your extensions stay active no matter how you launch Chrome (Taskbar, Start Menu, or clicking a link in another app).
* **The Runner-Up (Shortcut):** This only works if you launch Chrome specifically from *that* modified icon. If Chrome is opened by another program, the bypass won't load.

#### **Are both required?**
* **No, but they work better together.** 
Using the **Registry Fix** ensures the policy allows the extension to exist, while the **Shortcut Parameter** tells the Chrome engine to ignore the "Unsupported" warnings.

### 5. 📥 Extensions Installation
Open the generated **`Extensions.html`** in Chrome to begin restoring your setup:
*   **Option A (Direct Store Access):** 
    *   Click **View in Store** to jump directly to the official extension page using its unique ID. 
    *   If the ID link is dead (404), click **Search in Store** to find the latest version or a replacement by name.
*   **Option B (Sideloading):** For removed or legacy extensions, click the **Extension Name** to open its local folder, then **Drag & Drop** that folder into `chrome://extensions` (ensure **Developer Mode** is ON).

> **💡 POWER-USER TIP:** If your browser security blocks the folder from opening automatically when you click the name, don't sweat it! Just manually open the **Chrome Extensions - REMAS** folder on your desktop. From there, you can manually drag and load all your unavailable or older Manifest V2 extensions to bypass Google's restrictions like a boss.

---

## 📝 Release Notes - v0.79 (Beta)
> *"The Precision Update"*

*   **✨ New! 3-Label Badge System:** Every extension now displays:
    *   `Source Count` (One Source vs Both Sources).
    *   `Winner Origin` (Source 1 vs Source 2 - identifying which file was kept).
    *   `Version Badge` (Live version number audit).
*   **🛠 Smart Search Logic:** Fixed the `/search/` URL structure to find re-uploaded or rebranded extensions in the Store.
*   **📊 Optimized Header:** Added a real-time Audit Summary showing exactly how many duplicates were purged from each source.
*   **📂 Folder Collision Fix:** Added the `_d` logic. If two *different* extensions have the same name, REMAS now saves both instead of overwriting.
*   **🌍 Localized Recovery:** Improved "Deep Scan" for localized apps (Fixes the `__MSG_` naming bug).
*   **🔢 Sequential Numbering:** Added list numbering to the HTML for easier "check-listing" during bulk installs.

---

## 📝 Release Notes - v0.80 (Beta)
> *"The Precision Audit Update"*

*   **🧠 Intelligent Folder Logic:** New "All-or-Nothing" fallback system. If `Extensions_1` is missing on Desktop, the engine skips the desktop search to prevent errors and targets your live **Chrome System Path**.
*   **📂 Single-Source Mode:** The engine now intelligently toggles to a dedicated "Single-Source" mode during live fallbacks. This ensures a clean audit of your current active profile.
*   **🎯 View in Store Button:** Added a direct **ID-based** link next to every extension. Users can now jump directly to the specific official app page on the [Chrome Web Store](https://chrome.google.com).
*   **⚔️ Dual-Action Strategy:** Every entry now features both **View** and **Search** buttons. This "Dual-Store" approach ensures that even if an Extension ID is nuked (404), the name-based search fallback will find the replacement.
*   **🚀 Header Branding:** Updated the dashboard with a **Rocket icon** and version badge. Added a direct, styled link to the official [GitHub/REMAS](https://github.com/ZulfekarAliAgha/REMAS) repository.
*   **👤 Audit Highlighting:** The header now explicitly displays the **Windows Username** and **Total Extension Count** in **Bold Red** to provide instant visual confirmation of the audit scope.

---

