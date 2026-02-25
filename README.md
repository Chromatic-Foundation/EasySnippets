
# Easy Snippets for Word 2019

**Easy Snippets** is a Word 2019 add-in that lets you:

- Add custom text snippets (Add Snippets)  
- Search and insert snippets quickly (Show Snippets)  
- Automatically save snippets in `%appdata%\Microsoft\Word\EasySnippets\EasySnippets.txt`  

This add-in creates a **new Ribbon tab** called `Easy Snippets` that works in **all Word documents**, not just the `.dotm` file.

---

## Installation

### Option 1 — Load for all documents using Word Startup folder

1. Download the `EasySnippets.dotm` file from this repository.  
2. Locate Word's **Startup folder**:  
C:\Users<YourUser>\AppData\Roaming\Microsoft\Word\STARTUP
> Replace `<YourUser>` with your Windows username.

3. Move or copy `EasySnippets.dotm` into the **Startup** folder.  
4. Open Word 2019. The **Easy Snippets** tab will appear automatically in the Ribbon for **all documents**, new and existing.  

### Option 2 — Attach template to a specific document

1. Open Word 2019.  
2. Go to **File → Options → Add-Ins**.  
3. At the bottom, next to "Manage:", select **Templates** and click **Go...**.  
4. Click **Attach...** and select `EasySnippets.dotm`.  
5. Make sure **"Automatically update document styles"** is checked, then click **OK**.  
6. The **Easy Snippets** tab will appear for this document.

---

## Usage

### Add Snippets

1. Click **Add Snippets** in the Ribbon.  
2. Enter a **Snippet Name** and **Snippet Text**.  
3. Click **Save**.  
4. The snippet is now saved and available for future sessions.

### Show Snippets

1. Click **Show Snippets** in the Ribbon.  
2. Type the snippet name in the **search box**.  
3. Select the snippet from the list and click **Insert**.  
4. The snippet text will be inserted into the current document.

---

## Notes

- Snippets are saved automatically in `%appdata%\Microsoft\Word\EasySnippets\EasySnippets.txt`.  
- Once the `.dotm` is in the **Startup folder**, the tab is available for **all Word documents**.  
- This works in Word 2019 on Windows. Other versions may work but are untested.


