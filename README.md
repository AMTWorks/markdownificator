# Markdownificator

**English** · [Español](README.es.md)

Markdownificator v1.2 exports the document you have open in Word for Windows to Markdown, including unsaved changes. The conversion runs on your computer and produces a `.md` file in GitHub Flavored Markdown (GFM).

## Before you start

You need **desktop Word for Windows** and **Pandoc**, the tool that converts the document to Markdown. Markdownificator v1.2 was tested with Word for Microsoft 365, Windows 11, and Pandoc 3.14. The macro displays messages in Spanish when Word's menus are in Spanish and in English for all other Word interface languages.

## 1. Install Pandoc

If you already have Pandoc, open **Terminal** or **Command Prompt** from the Start menu and run `pandoc --version`. If you see `pandoc` followed by a version number, go to step 2.

To install it, close Word, open **Terminal** or **Command Prompt**, paste the following command, and press **Enter**. Follow the prompts.

```powershell
winget install --source winget --exact --id JohnMacFarlane.Pandoc
```

If Windows doesn't recognize `winget` or you'd prefer a setup wizard, download and run the **Windows installer (`.msi`)** from Pandoc's official site. With either method, open a new terminal window after installation and run `pandoc --version` to check that Windows can find the program. [Official Pandoc installation guide](https://pandoc.org/installing.html#windows)

If PowerShell still doesn't recognize `pandoc` in the window where you installed it with WinGet, close that window, open a new one, and run `pandoc --version` again before reinstalling anything. The old window may still be using the settings it had before you installed Pandoc.

## 2. Download and import the macro

1. Open [Markdownificator_v1.2.bas](Markdownificator_v1.2.bas) on GitHub. Click the download icon in the upper-right corner of the file view (*Download raw file*) and save the `.bas` file to **Downloads**. [Download a file from GitHub](https://docs.github.com/en/repositories/working-with-files/using-files/viewing-and-understanding-files#viewing-or-copying-the-raw-file-content)
2. Open Word and create a blank document. Press **Alt + F11** to open the Visual Basic Editor. You can also go to **Developer → Visual Basic**. If the **Developer** tab is hidden, go to **File → Options → Customize Ribbon**, check **Developer** under **Main Tabs**, and click **OK**. [Show the Developer tab](https://support.microsoft.com/en-us/word/show-the-developer-tab-in-word)
3. In the editor's left pane, click the **Normal** project. If the pane is hidden, press **Ctrl + R**. **Normal** is the personal template Word already uses. Saving the macro there makes it available in other documents. [Project Explorer](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/view-menu) · [Make macros available in all documents](https://support.microsoft.com/en-us/word/create-or-run-a-macro)
4. In the editor, go to **File → Import File**, select `Markdownificator_v1.2.bas`, and click **Open**. Check that **Normal → Modules → Markdownificator_v1_2** appears in the pane. [Import a module in Visual Basic](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-menu)
5. With **Normal** selected, click the editor's **Save** icon. Press **Alt + F11** to return to Word, then close Word. If Word asks whether to save changes to the **Normal** template, choose **Save**.

These steps are for a first installation. If you've already installed an earlier or test version of Markdownificator, check the full macro name when adding the button. Older versions may also appear as `ExportarAMarkdown`. If you already see `Markdownificator_v1_2` under **Normal → Modules**, check which version you have before importing again. Word can append a number to the name of a duplicate module. [How Word imports duplicate modules](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-menu)

## 3. Add the Export to Markdown button

1. Reopen Word and go to **File → Options → Quick Access Toolbar**.
2. Under **Choose commands from**, select **Macros**. Look for `Normal.Markdownificator_v1_2.ExportarAMarkdown`. Word may shorten the displayed name.
3. In the drop-down on the right, select **For all documents (default)**. Select the macro on the left and click **Add**.
4. Select the macro you added on the right, click **Modify**, pick an icon, and enter **Export to Markdown** as the display name. Click **OK** to close each dialog. [Assign a button to a macro](https://support.microsoft.com/en-us/word/create-or-run-a-macro)

If the toolbar is hidden, right-click the ribbon and select **Show Quick Access Toolbar**. You can also run the macro from **View → Macros → View Macros**. [Show the Quick Access Toolbar](https://support.microsoft.com/en-us/office/customize-the-quick-access-toolbar) · [Run a macro](https://support.microsoft.com/en-us/word/create-or-run-a-macro)

## 4. Export your first document

1. Open a document in Word. For a quick test, create a blank document and type a line of text.
2. Click **Export to Markdown** on the Quick Access Toolbar. A save dialog should open with **Markdownificator v1.2** in its title.
3. Choose where to save the file, click **Save**, then open the `.md` file in Notepad to check that it contains your text. If the document contains pictures, the macro may warn you that the Markdown keeps the image references while the image files themselves are left out of the export.

Exporting the image files is planned for a future version.

If the macro reports that Pandoc is missing, run `pandoc --version` in a new terminal window. If the command works, close all Word windows and reopen Word before trying the export again. If Word says your organization blocks macros, contact your IT support team. [Macro settings in Office](https://support.microsoft.com/en-us/office/vba/enable-or-disable-macros-in-microsoft-365-files)

## Where Markdown files are saved

Each time you export, the macro asks Windows where your account's **Downloads** folder is. If it finds the folder, it suggests it in the save dialog even if you've moved Downloads to another drive or saved the previous export elsewhere. You can change the folder and filename before clicking **Save**.

If Windows doesn't return a valid path or the macro cannot confirm that the folder exists, the macro suggests only the filename and Windows chooses which folder to open. Check the location in the save dialog and choose the folder you want. This behavior was simulated using the Windows save dialog. It still needs to be checked in Word using an account whose Downloads folder is inaccessible.

## Help and bug reports

If you run into a problem, open an issue in the repository's **Issues** tab. Include your Windows, Word, and Pandoc versions, the step where the problem occurred, and any error message. If the save dialog opens in another folder, mention whether you can open **Downloads** in File Explorer and describe the folder that appeared (for example, “the document's folder” or “the last folder I used”). Remove any personal details from error messages or paths. A description of what happened is enough. [Open an issue on GitHub](https://docs.github.com/en/issues/tracking-your-work-with-issues/using-issues/creating-an-issue)

## Credits

- **J.F. Franco (@AMTWorks)** · Project direction and testing in Word.
- **Claude (Anthropic)** · Development of v1.1.
- **ChatGPT (OpenAI)** · Development of v1.1.1 and v1.2.

## License

Markdownificator is available under the [MIT License](LICENSE).
