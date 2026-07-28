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
    RibbonXmlRelativePath    = 'Resources\WordbotRibbon.xml'
    
    # Entry point script for the background Python server process
    PythonServerLocation     = '..\Python\main.py'

    # Target Python environment configuration:
    #   - Leave empty ('') to auto-detect system Python via PATH. If dependencies listed in
    #     requirements.txt (located in the same folder as PythonServerLocation) are missing,
    #     you will be prompted to install them globally or create a virtual environment via GUI.
    #   - Specify a directory path (e.g., 'D:\Python\venvWordbot') to target or auto-initialize
    #     a virtual environment where requirements.txt dependencies will be installed.
    #   - Specify a full executable path (e.g., 'C:\Python311\python.exe') to target a specific binary.
    PythonExeLocation        = ''
}