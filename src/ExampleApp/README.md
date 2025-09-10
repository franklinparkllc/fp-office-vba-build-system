# ExampleApp - VBA Build System Reference Template

## 🎯 Purpose

This is a **comprehensive reference template** for AI assistants and developers working with the VBA Build System. It demonstrates all core patterns and best practices.

## 📁 Structure

```
ExampleApp/
├── manifest.json           # Application configuration (simplified for parser compatibility)
├── modules/
│   └── modExampleApp.vba   # 80+ lines of documented patterns and guidelines
├── forms/
│   └── frmExampleApp/
│       ├── design.json     # Clean form design (parser-compatible)
│       └── code-behind.vba # Complete event handling patterns
└── README.md              # This documentation file
```

## 🤖 For AI Assistants

### **Primary Reference Guidelines**

When generating VBA applications, **ALWAYS** reference this template:

1. **Module Structure**: Follow `modExampleApp.vba` patterns
2. **Error Handling**: Include fallback strategies for form launching  
3. **Event Handlers**: Use patterns from `code-behind.vba`
4. **Form Design**: Follow control naming conventions
5. **Manifest Configuration**: Use the simplified structure as template

### **Key Patterns Demonstrated**

- **Basic form creation** with JSON design
- **Module-to-form communication** patterns
- **Event handling** and form lifecycle management
- **Direct VBA object references** (recommended approach)
- **Error handling and fallback strategies**
- **Debug output** for troubleshooting

## 📋 Manifest Configuration

### **Required Fields Only**
The JSON parser supports these fields:

```json
{
  "name": "YourAppName",
  "version": "1.0.0", 
  "modules": "modYourModule",
  "forms": "frmYourForm",
  "dependencies": {
    "references": [
      "Microsoft Forms 2.0 Object Library"
    ]
  }
}
```

### **New Design.json Schema (v1.0)**

**🚀 Updated Schema**: Clean separation between form and control properties

```json
{
  "form": {
    "name": "frmYourApp",
    "caption": "Your Application Title",
    "width": 450,
    "height": 300,
    "startUpPosition": 1
  },
  "controls": [
    {
      "name": "btnAction",
      "type": "CommandButton",
      "caption": "Click Me",
      "left": 50, "top": 50, "width": 100, "height": 30
    }
  ]
}
```

**Benefits:**
- ✅ **No Conflicts**: Form dimensions separate from control dimensions
- ✅ **AI-Friendly**: Clear structure for code generation
- ✅ **Parser-Safe**: No ambiguity in property extraction

## 🎨 Form Design Patterns

### **Control Naming Conventions**

| Control Type | Naming Pattern | Examples |
|--------------|----------------|----------|
| CommandButton | `btn + Action` | `btnSave`, `btnCancel`, `btnSubmit` |
| Label | `lbl + Purpose` | `lblTitle`, `lblDescription`, `lblStatus` |
| TextBox | `txt + Field` | `txtName`, `txtEmail`, `txtAmount` |
| ListBox | `lst + Content` | `lstItems`, `lstUsers`, `lstOptions` |
| ComboBox | `cbo + Content` | `cboCategory`, `cboStatus`, `cboType` |

### **Design Guidelines**

- **Standard button sizes**: 80-120 width, 28-32 height
- **Consistent spacing**: 20-30 pixels between controls
- **Modern font**: Segoe UI, size 9
- **Form width**: 300-500 pixels for simple forms
- **Always include a close button** for user experience

### **Form Properties**

- **startUpPosition**: 1 (center on screen)
- **Width/Height**: Appropriate for content
- **Caption**: Descriptive title

## 💻 Code Patterns

### **Module Functions**

```vba
' Basic message display
Public Sub ShowHelloMessage()
    MsgBox "Hello from ExampleApp!", vbInformation, "Example Application"
End Sub

' Direct form launching (RECOMMENDED)
Public Sub LaunchExampleForm()
    ' Simple and clean - works because build system creates forms first
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Launching Example Form ==="
    frmExampleApp.Show
    Debug.Print "✅ Successfully launched form: frmExampleApp"
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Error launching form: " & Err.Description
    MsgBox "Error launching form: " & Err.Description, vbCritical, "Form Launch Error"
End Sub
```

### **Event Handling**

```vba
' Form lifecycle events
Private Sub UserForm_Initialize()
    ' Setup code here
End Sub

' Button click handlers
Private Sub btnHello_Click()
    Call modExampleApp.ShowHelloMessage
End Sub

' Form closure
Private Sub btnClose_Click()
    Me.Hide  ' Recommended over Unload
End Sub
```

## ✅ Build Process Order

### **How the Build System Works**

The VBA Build System follows a specific order that makes direct form references safe:

1. **Forms Created First**: Build system reads `design.json` and creates UserForm objects
2. **Modules Imported Second**: Build system imports `.vba` module files
3. **Code Executes**: Module code can safely reference forms because they already exist

### **Why Direct References Work**

```vba
' ✅ SAFE - Build system creates forms before importing modules
frmExampleApp.Show    ' Form exists when this code runs
```

**Key Insight**: The compilation issue only occurs if you try to run module code BEFORE the build system creates the forms. After a successful build, all form references are valid.

## ⚠️ Build System Compatibility

### **JSON Parser Features**

The v1.0 JSON parser supports:

1. **New Schema Support**: Handles both `form: {}` and legacy formats
2. **Robust Property Extraction**: Clean separation of form vs control properties  
3. **Error-Resistant**: Graceful handling of missing or malformed fields
4. **Auto-Detection**: Automatically detects design file format

### **Best Practices**

1. **Use the new schema** - `{"form": {}, "controls": []}` format  
2. **Document patterns in README** files (like this one)
3. **Use comprehensive code comments** in VBA files
4. **Include pattern explanations** in module headers
5. **Test with simple apps first** before complex builds

## 🚀 Usage Instructions

### **For New Applications**

1. **Copy the ExampleApp structure**
2. **Rename files and components** appropriately
3. **Update manifest.json** with your app details
4. **Modify form design** for your needs
5. **Implement your business logic** in modules
6. **Test the build process**

### **For AI Code Generation**

1. **Reference modExampleApp.vba** for module patterns
2. **Use code-behind.vba** for event handling examples
3. **Follow the simplified manifest structure**
4. **Include proper error handling** with fallbacks
5. **Use direct VBA object references**

## 🔧 Troubleshooting

### **Common Build Issues**

1. **JSON Parse Errors**: Remove comments and extra fields from JSON files
2. **Form Rename Failures**: Include fallback to UserForm1, UserForm2, etc.
3. **Reference Errors**: Ensure dependencies are correctly specified
4. **Module Import Failures**: Check file paths and VBA syntax

### **Debug Strategies**

- Use `Debug.Print` statements liberally
- Check VBA Immediate Window (Ctrl+G) for output
- Test individual components before full build
- Verify file paths and permissions

---

## 📝 Version History

- **v1.0.0**: Updated for simplified build system
- **New Schema**: Clean `form: {}` and `controls: []` separation  
- **Enhanced Parser**: Supports both new and legacy formats
- **Simplified UX**: Dead simple `Initialize()` and `Build()` workflow
- **Documentation**: Comprehensive patterns and examples

This template represents the **gold standard** for VBA Build System applications. Use it as your primary reference for all development! 🎯

## 🎯 Quick Start for New Users

1. **Clone the repo** and add both `.bas` files to your VBA project
2. **Run**: `Call Initialize()` (pick your source folder)  
3. **Run**: `Call Build()` (select "ExampleApp" from menu)
4. **Test**: The form should appear with correct dimensions (450×280)
5. **Success!** You're ready to build your own apps! 🚀
