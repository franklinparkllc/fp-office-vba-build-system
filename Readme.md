# VBA Builder - A Modern Build System for Microsoft Office
## 🚀 Overview
Modern build system for VBA development, optimized for AI-driven workflows. This system transforms VBA development by enabling version control, automated builds, and modern development practices with minimal complexity.
## ✨ What Makes This Special
### **Before VBA Builder**
- ❌ Code trapped inside Office documents
- ❌ No version control for VBA code
- ❌ Manual, error-prone form creation and management
- ❌ No automated deployment process

### **After VBA Builder**
- ✅ **Source-Controlled Code**: All VBA code lives in text files, perfect for Git.
- ✅ **Automated Form Generation**: Build complex UserForms directly from `design.json` files.
- ✅ **Minimal Dependencies**: The entire build system is self-contained in a single VBA module.
- ✅ **AI-Optimized**: Designed for agentic workflows, enabling AI to generate, build, and test applications.
- ✅ **Radical Simplicity**: 80% less code than traditional build systems, but with the same power.

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
## 🏗️ **System Architecture (v1.0 - Simplified)**
**Two-module design**: Dead simple user experience with powerful features under the hood.

### **👤 User Experience - Just 2 Functions!**
```vba
Call Initialize()    ' Setup (one-time, opens folder picker)
Call Build()         ' Build apps (shows menu, auto-initializes)
```

### **modBuildSystem.bas** - Core Build Engine
- **🎯 User Functions**:
  - `Initialize()` - One-time setup with folder picker
  - `Build()` - Shows available apps, builds selected one
  - `BuildApplication(appName)` - Direct build for specific app
  - `ShowSystemStatus()` - Check current configuration
- **🔧 Internal Engine**:
  - `CreateFormDirect()` - Robust form creation with new schema support
  - `ProcessModules()` - Module importing
  - `ApplyDesign()` - Form design application with improved JSON parsing

### **modBuilderUtils.bas** - Utility Functions
- **File & Project Management**: Safe file I/O, VBA project manipulation
- **Configuration Management**: Persistent settings, path validation
- **System Validation**: Trust settings, folder scanning, error handling

### **🚀 Key Features**:
- ✅ **Dead Simple**: 2-function user experience
- ✅ **Auto-Initialization**: Build functions set up automatically
- ✅ **Persistent Settings**: Set source path once, works forever
- ✅ **New JSON Schema**: Clean separation of form properties and controls
- ✅ **Folder Picker**: No typing file paths
## 📁 **Project Structure**
```
YourProject/
├── modBuildSystem.bas     # Core build engine (v1.0)
├── modBuilderUtils.bas    # Utility functions
└── src/                   # Your application source files
    ├── ExampleApp/        # A reference application
    │   ├── manifest.json  # App config: name, modules, forms
    │   ├── modules/
    │   │   └── modExampleApp.vba
    │   └── forms/
    │       └── frmExampleApp/
    │           ├── design.json      # Form layout: size, caption, controls
    │           └── code-behind.vba  # The form's event-handling code
    └── YourApp/           # Your new application
        ├── manifest.json
        ├── modules/
        └── forms/
```
## 🚀 **Quick Start**
### **1. Setup VBA Builder**
1.  **Open any Office document** (Excel, Word, etc.) and save it as a macro-enabled file (e.g., `.docm` or `.xlsm`). This will be your build environment.
2.  **Press `Alt+F11`** to open the VBA Editor.
3.  **Insert → Module** and name it `modBuildSystem`.
4.  **Copy the entire contents** of `modBuildSystem.bas` into the module.
5.  **Insert → Module** and name it `modBuilderUtils`.
6.  **Copy the entire contents** of `modBuilderUtils.bas` into the second module.
### **2. Enable Trust Center Settings**
**CRITICAL**: You must allow programmatic access to the VBA project.
1.  In your Office Application: **File → Options → Trust Center**.
2.  **Trust Center Settings → Macro Settings**.
3.  **✅ Check "Trust access to the VBA project object model"**.
4.  **Restart the Office application**.
### **3. Start Building! (Dead Simple)**
```vba
' Open the Immediate Window (View -> Immediate Window or Ctrl+G) and run:

' First time setup (opens folder picker)
Call Initialize()

' Build apps (shows menu, auto-setup if needed)
Call Build()

' Optional: Direct build
Call BuildApplication("ExampleApp")

' Optional: Check status
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
### **New Design.json Schema (v1.0)**
Clean separation between form properties and controls.
```json
{
  "form": {
    "name": "frmMyApp",
    "caption": "My Application", 
    "width": 400,
    "height": 300,
    "startUpPosition": 1
  },
  "controls": [
    {
      "name": "btnSubmit",
      "type": "CommandButton", 
      "caption": "Submit",
      "left": 50, "top": 50, "width": 100, "height": 30
    }
  ]
}
```

**🎯 Why the new schema?**
- ✅ **No Conflicts**: Form width/height separate from control dimensions
- ✅ **AI-Friendly**: Clear structure for code generation  
- ✅ **Future-Proof**: Easy to add form-level properties
- ✅ **Self-Documenting**: Obvious what belongs where
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
- **Location**: `src/ExampleApp/`
### **FundSelector** - Production Example
- **Purpose**: A database-driven fund selection tool.
- **Features**: Demonstrates Azure SQL connectivity and a professional UI.
- **Use Case**: A real-world business application.
## 🔧 **System Commands**
```vba
' 🎯 Essential (98% of use cases)
Call Initialize()                    ' Setup/change source folder (folder picker)
Call Build()                         ' Show menu, build selected app

' 🔧 Direct commands  
Call BuildApplication("AppName")     ' Build specific app directly
Call ShowSystemStatus()              ' Check current configuration

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
' Step 4: Re-setup if needed
Call Initialize()
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
### **For AI Assistants: Simplified Code Generation**
-   **Simple JSON Structures**: Easy for AI to generate `design.json` files.
-   **Direct Object References**: AI can generate code that uses `frmYourForm.Show` syntax, as the build process guarantees the form will exist.
-   **Modular & Reliable**: The two-module build system is well-organized with clear separation of concerns.
---
## 🎉 **Ready for Modern VBA Development?**
The **simplified v1.0 system** is perfect for everyone - from beginners to AI assistants. Just 2 commands to get started!

```vba
' Dead simple workflow:
Call Initialize()    ' Pick your source folder (one-time)
Call Build()         ' Build your apps (menu-driven)
```

**Perfect for:**
- 🤖 **AI Code Generation** - Clean schema and simple commands
- 👨‍💻 **Developers** - Modern workflow with version control
- 🏢 **Organizations** - Standardized, repeatable builds
- 📚 **Learning** - Clear patterns and comprehensive examples