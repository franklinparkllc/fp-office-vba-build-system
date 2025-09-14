# VBA App Builder - A Modern Build System for Microsoft Office
## 🚀 Overview
Modern build system for VBA development, optimized for AI-driven workflows. This system transforms VBA development by enabling version control, automated builds, and modern development practices with minimal complexity.
## ✨ What Makes This Special
### **Before VBA App Builder**
- ❌ Code trapped inside Office documents
- ❌ No version control for VBA code
- ❌ Manual, error-prone form creation and management
- ❌ No automated deployment process

### **After VBA App Builder**
- ✅ **Source-Controlled Code**: All VBA code lives in text files, perfect for Git.
- ✅ **Automated Form Generation**: Build complex UserForms directly from `design.json` files.
- ✅ **Zero Configuration**: No persistent settings - select your source folder each time you build.
- ✅ **Single Module Simplicity**: The entire build system in one self-contained VBA module.
- ✅ **AI-Optimized**: Designed for agentic workflows, enabling AI to generate, build, and test applications.
- ✅ **Enhanced UX**: Progress tracking, better error messages, and auto-save options.

## 🔧 How It Works: The Direct Injection Method
The build system works by directly manipulating the VBA project within the host Office application (Word, Excel, etc.). It reads source files from your `src` directory and injects them into the VBA environment.

### 1. **Direct Form Creation**
Instead of relying on a fragile export/import process, the builder creates forms from scratch, overcoming wellknown limitations of the VBA environment.
```vba
' Create a new, empty form component
Set formComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
' Apply all properties from design.json (caption, size, etc.)
Call ApplyDesign(formComp, design)
' PAUSE to allow the VBE to process the new object
DoEvents: Sleep 500
' Programmatically rename the form to its correct name
formComp.Properties("Name").Value = formName
```
### 2. **Simple JSON Parsing**
A lightweight, regex-based parser reads `manifest.json` and `design.json` files. It's fast, reliable, and has no external dependencies.
```vba
' Extract all key/value pairs from a JSON file
Set dict = ParseJSON(jsonText)
' Access values directly
formComp.Properties("Width").Value = dict("width")
```
### 3. **Module & Code-Behind Injection**
Standard modules (`.vba`) are imported directly, and form code-behinds are injected into the form's `CodeModule`.
```vba
' Import a standard module from a .vba file
Set comp = vbProj.VBComponents.Import(filePath)
' Add event handlers and logic to a form
formComp.CodeModule.AddFromString codeContent
```
## 🏗️ **System Architecture (v2.1 - Simplified Edition)**
**Single-module design**: Everything you need in one file with zero configuration.

### **👤 User Experience - Just 2 Main Functions!**
```vba
Call Build()                        ' Browse for folder, select app, build
Call BuildApplication("AppName")    ' Browse for folder, build specific app
```

### **modAppBuilder.bas** - Complete Build System
- **🎯 User Functions**:
  - `Build()` - Browse for source folder and select app to build
  - `BuildApplication(appName)` - Build specific app (prompts for folder)
  - `ConfigureAutoSave()` - Toggle auto-save preference
  - `ShowSystemStatus()` - Display version and available commands
- **🔧 Core Features**:
  - `CreateFormDirect()` - Robust form creation with enhanced error handling
  - `ProcessModules()` - Module importing with progress tracking
  - `JSON Comment Support` - Allows // comments in JSON files
  - `JSON Validation` - Reports line numbers for syntax errors
  - `Progress Tracking` - Real-time build progress in Immediate window

### **🚀 Key Features**:
- ✅ **Zero Configuration**: No setup required - just import and use
- ✅ **Single Module**: All functionality in one easy-to-distribute file
- ✅ **Fresh Folder Selection**: Choose your project folder each time
- ✅ **Enhanced Error Messages**: Helpful suggestions for common issues
- ✅ **Progress Tracking**: See exactly what's happening during builds
- ✅ **JSON Comments**: Add // comments to your JSON files
- ✅ **Auto-Save Option**: Optionally save project after successful builds
## 📁 **Project Structure**
```
fp-office-vba-build-system/
├── modAppBuilder.bas      # Complete build system (v2.1)
├── ExampleApp/            # Reference application (at root for easy access)
│   ├── manifest.json      # App config: name, modules, forms
│   ├── modules/
│   │   └── modExampleApp.vba
│   └── forms/
│       └── frmExampleApp/
│           ├── design.json      # Form layout: size, caption, controls
│           └── code-behind.vba  # The form's event-handling code
└── YourAppsFolder/        # Your applications go here
    └── YourApp/
        ├── manifest.json
        ├── modules/
        └── forms/
```
## 🚀 **Quick Start**
### **1. Setup VBA App Builder**
1.  **Open any Office document** (Excel, Word, etc.) and save it as a macro-enabled file (e.g., `.docm` or `.xlsm`). This will be your build environment.
2.  **Press `Alt+F11`** to open the VBA Editor.
3.  **Insert → Module** and name it `modAppBuilder`.
4.  **Copy the entire contents** of `modAppBuilder.bas` into the module.
### **2. Enable Trust Center Settings**
**CRITICAL**: You must allow programmatic access to the VBA project.
1.  In your Office Application: **File → Options → Trust Center**.
2.  **Trust Center Settings → Macro Settings**.
3.  **✅ Check "Trust access to the VBA project object model"**.
4.  **Restart the Office application**.
### **3. Start Building! (Zero Configuration)**
```vba
' Open the Immediate Window (View -> Immediate Window or Ctrl+G) and run:

' Build apps (browse for folder, select app)
Call Build()

' Or build specific app directly
Call BuildApplication("ExampleApp")

' Optional: Configure auto-save
Call ConfigureAutoSave()

' Optional: Check version and commands
Call ShowSystemStatus()
```
## 📋 **Application Configuration**
### **Simplified Manifest.json**
This file defines your application's components.
```json
{
  "name": "MyApp",
  "version": "1.0.0",
  "modules": "modMyModule",
  "forms": "frmMyForm"
}
```
### **Design.json Schema (with Comments!)**
Clean separation between form properties and controls, now with comment support.
```json
{
  // Form properties section
  "form": {
    "name": "frmMyApp",
    "caption": "My Application", 
    "width": 400,
    "height": 300,
    "startUpPosition": 1  // 0=Manual, 1=CenterOwner, 2=CenterScreen, 3=WindowsDefault
  },
  // Controls array - add buttons, labels, etc.
  "controls": [
    {
      "name": "btnSubmit",
      "type": "CommandButton", 
      "caption": "Submit",
      "left": 50, "top": 50, "width": 100, "height": 30,
      // Optional font settings
      "font": {
        "name": "Arial",
        "size": 10,
        "bold": true,
        "italic": false
      }
    }
  ]
}
```

**🎯 Enhanced Features:**
- ✅ **JSON Comments**: Use // for single-line comments
- ✅ **Font Support**: Customize control fonts
- ✅ **Clear Structure**: Form properties separate from controls
- ✅ **AI-Friendly**: Easy for code generation
- ✅ **Error Reporting**: Get line numbers for JSON syntax errors
## 🎯 **The Direct VBA Object Strategy**
After a successful build, your forms and modules exist as **native VBA objects**. You should always reference them directly in your code.
### ✅ **Recommended Approach**
```vba
' Launch forms using their given name
frmMyForm.Show
' Call module functions directly
Call modMyModule.MyFunction()
' Access form properties directly
frmMyForm.Caption = "New Title"
```
### 🚫 Critical Rule: Never reference placeholder names like "UserForm1"
- Always assume build-created objects exist with their manifest/design names.
- The build process ensures these objects are created before your code runs. Do not write code that references ephemeral placeholder names such as `UserForm1`, as they are not guaranteed to exist.
## 📚 **Reference Applications**
### **ExampleApp** - Reference Template 🎯
- **Purpose**: A comprehensive template for developers and AI assistants.
- **Features**: A fully annotated application demonstrating all VBA Build System patterns.
- **Use Case**: Copy this structure when creating new applications.
- **Location**: `ExampleApp/` (at root level for easy access)
### **FundSelector** - Production Example
- **Purpose**: A database-driven fund selection tool.
- **Features**: Demonstrates Azure SQL connectivity and a professional UI.
- **Use Case**: A real-world business application.
## 🔧 **System Commands**
```vba
' 🎯 Essential Commands
Call Build()                         ' Browse for folder, show menu, build selected app
Call BuildApplication("AppName")     ' Browse for folder, build specific app

' 🔧 Configuration & Info
Call ConfigureAutoSave()             ' Toggle auto-save on/off
Call ShowSystemStatus()              ' Display version and available commands

' 🔍 Diagnostics (if needed)
Call ValidateTrustSettings()         ' Check VBA project access
```
## 🛠️ **The Form Build Process Explained**
The system uses a robust, multi-step process to overcome the quirks of the VBE and reliably generate forms.
1.  **Create**: A new, blank `MSForm` component is added to the project with a temporary name (e.g., `UserForm1`).
2.  **Apply Design**: The system reads your `design.json` and programmatically applies all properties (caption, size, etc.) and adds all specified controls to the new form.
3.  **Pause**: A brief, 500ms pause allows the VBE to finish processing the new form and its controls, preventing race conditions.
4.  **Rename**: With the form fully created, the system renames it from its temporary name to the name specified in your configuration (e.g., `frmExampleApp`).
5.  **Inject Code**: The associated `code-behind.vba` file is read and injected into the form's code module.
6.  **Save**: The host document is saved to persist all changes to the VBA project.
## 🚨 **Troubleshooting**
### **Common Issues**
1.  **"VBA Project access is disabled"**
    -   ✅ Enable "Trust access to the VBA project object model" in Trust Center and restart.
2.  **"Build failed during processing"**
    -   ✅ Check the file paths shown in `ShowSystemStatus()`.
    -   ✅ Verify the syntax of your `manifest.json` and `design.json` files.
3.  **"Application not found"**
    -   ✅ Check `ShowSystemStatus()` to see available applications.
    -   ✅ Ensure your application folder in `src` contains a valid `manifest.json`.
### **Diagnostic Steps**
```vba
' Step 1: Check system status
Call ShowSystemStatus()
' Step 2: Validate Trust Center settings  
Call ValidateTrustSettings()
' Step 3: Try building with menu
Call Build()
' Step 4: Check auto-save settings
Call ConfigureAutoSave()
```
## 🏆 **Benefits**
### **For Developers**
-   **Modern Workflow**: Develop VBA using modern IDEs (VS Code, Cursor).
-   **Version Control**: Full Git integration for all your code.
-   **Collaboration**: Team development with proper source control.
-   **Native Experience**: Debugging and IntelliSense work perfectly.
### **For Organizations**
-   **Standardization**: Consistent, repeatable builds and deployments.
-   **Quality Control**: Automated build validation.
-   **Maintainability**: Easier updates and bug fixes.
### **For AI Assistants: Enhanced Code Generation**
-   **Comment-Friendly JSON**: Generate design files with helpful // comments.
-   **Better Error Messages**: Line numbers in JSON errors make debugging easier.
-   **Progress Tracking**: Monitor build progress programmatically.
-   **Single Module**: Everything in one file - simpler to understand and generate for.
---
## 🎉 **Ready for Modern VBA Development?**
The **v2.1 system** brings zero-configuration simplicity with enhanced user experience!

```vba
' Ultra-simple workflow:
Call Build()         ' Browse, select, build - all in one!
```

**What's New in v2.1:**
- 🚀 **Zero Configuration** - No setup, no stored paths
- 📊 **Progress Tracking** - See build progress in real-time
- 💬 **JSON Comments** - Add documentation to your config files
- 🔍 **Better Errors** - Helpful messages with recovery suggestions
- 💾 **Auto-Save Option** - Save project after successful builds
- 📦 **Single Module** - Everything in one easy-to-share file

**Perfect for:**
- 🤖 **AI Code Generation** - Enhanced JSON support and error handling
- 👨‍💻 **Developers** - Modern workflow with better UX
- 🏢 **Organizations** - Zero-config deployment
- 📚 **Learning** - Clearer error messages and progress feedback