# Installation Guide - macOS

This guide covers installing Wordbot on macOS.

## Choose Your Edition

| Edition | What it does | Python Required | Zotero Required |
|---------|--------------|-----------------|-----------------|
| Markdown | Core Markdown formatting only | No | No |
| LLM | Markdown + AI model integration | Yes | No |
| Research | Markdown + AI + Zotero citations | Yes | Yes |

## Step 1: Download Pre-built Templates

The templates are located in `Powershell/build/`:
- `Wordbot_Markdown.dotm` - Core Markdown formatting features
- `Wordbot_LLM.dotm` - Markdown + LLM integration
- `Wordbot_Research.dotm` - Markdown + LLM + Zotero research features

## Step 2: Remove Quarantine Attribute

When you download a file from the internet, macOS may quarantine it and disable macros. Replace `Wordbot_Markdown.dotm` file name with your downloaded file name before running the following command:

```bash
xattr -d com.apple.quarantine Wordbot_Markdown.dotm
```


> [!IMPORTANT]
> If you skip this step, Word may not load the template or may block macros from running.

## Step 3: Set Up Python (LLM and Research Editions Only)

> [!NOTE]
> Skip this step if you are using `Wordbot_Markdown.dotm`.

Follow the instructions in [PythonGuide.md](PythonGuide.md) to set up Python and configure the LLM endpoint.

## Step 4: Set Up Zotero (Research Edition Only)

> [!NOTE]
> Skip this step if you are using `Wordbot_Markdown.dotm` or `Wordbot_LLM.dotm`.

Follow the instructions in [ZoteroGuide.md](ZoteroGuide.md) to set up Zotero and Zotseek.

## Step 5: Copy WordbotCurl.scpt

Wordbot requires an AppleScript helper to communicate with the Python server for LLM and Research editions on macOS.

Copy `WordbotCurl.scpt` (located in [../Extras/macOS/](../Extras/macOS/WordbotCurl.scpt)) to:

```
~/Library/Application Scripts/com.microsoft.Word/WordbotCurl.scpt
```

> [!IMPORTANT]
> If this script is not in the correct location, the **Start Server** button will not work on macOS.

## Step 6: Copy Template to Word STARTUP Folder

Copy your chosen `.dotm` file to the Word STARTUP folder:

**Terminal Path:**
```
~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word
```

**Finder Path (Cmd + Shift + G):**
```
~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word
```

> [!NOTE]
> macOS hides the `.localized` extensions in Finder, but the literal path with `.localized` is required when using Terminal commands.

## Step 7: Launch Word

Open Microsoft Word. You should see a **Wordbot** tab in the ribbon.

> [!NOTE]
> If the Wordbot tab does not appear, restart Word completely. Ensure the `.dotm` file is unblocked and in the correct STARTUP folder.

---

## Next Steps

See [WordbotUserGuide.md](WordbotUserGuide.md) for features and buttons.

## Troubleshooting

If you encounter issues, see [Troubleshooting.md](Troubleshooting.md) for common problems and solutions.