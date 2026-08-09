@{
    # The full template filename including extension
    TemplateName             = 'Wordbot.dotm'
    
    # Base filename for the macro-enabled file (without extension)
    FileName                 = 'Wordbot'
    
    # Supported macro file extensions:
    #   - '.dotm' for global Word templates (loads automatically on startup)
    #   - '.docm' for macro-enabled documents (testing/standalone)
    #   Note: .docx and .doc formats are strictly forbidden as they cannot store VBA macros
    Extension                = '.dotm'
    
    # Target save directory (absolute path) for the output file:
    #   - Leave empty ('') to automatically target Word's default STARTUP folder
    #   - Or specify an absolute folder path (e.g., 'D:\Worddocs')
    outputPath               = ''
    
    # Relative paths (resolved against the location of installWordBotTemplate.ps1):
    # Directory containing raw VBA source modules (.bas, .cls, .frm)
    VBAFolderRelativePath    = 'Resources\WordVBAmacros'

    # Custom Office UI Ribbon layout definition XML
    RibbonXmlRelativePath    = 'Resources\RibbonXML'
    
    # Entry point script for the background Python server process
    PythonServerLocation     = '..\Python\main.py'

    # Target Python environment configuration:
    #   Option 1: Auto-detect - uses Python from system PATH
    #     If dependencies are missing, you'll be prompted to install globally or create a venv via GUI
    #     PythonExeLocation = ''
    #
    #   Option 2: Virtual environment folder - uses or creates a venv at the specified directory
    #     Dependencies will be installed inside this isolated environment
    #     PythonExeLocation = 'D:\wordbotTemp'
    #
    #   Option 3: Specific Python executable - points directly to a Python interpreter
    #     Can be system Python (C:\Python39\python.exe) OR a venv's python.exe
    #     To use specific Python version with venv: create venv first, then point to its python.exe
    #     PythonExeLocation = 'C:\Python311\python.exe'
    #     PythonExeLocation = 'D:\myproject\venv\Scripts\python.exe'
    PythonExeLocation        = ''
}