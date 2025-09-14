# ExampleApp - Reference Template

A complete example demonstrating VBA App Builder patterns and best practices.

## 📁 Structure

```
ExampleApp/
├── manifest.json           # App configuration
├── modules/
│   └── modExampleApp.vba   # Module with documented patterns
└── forms/
    └── frmExampleApp/
        ├── design.json     # Form layout and controls
        └── code-behind.vba # Event handlers
```

## 🚀 Quick Start

1. Run `Build()` in VBA and select the folder containing this ExampleApp
2. The form should appear with correct dimensions (450×280)
3. Click buttons to test functionality

## 📋 Key Files

### manifest.json
Defines the application name, modules, and forms to build.

### design.json  
Uses the new schema with separate `form` and `controls` sections:
- Form properties: name, caption, dimensions, position
- Controls: buttons, labels with positioning
- Supports font customization and // comments

### modExampleApp.vba
Demonstrates:
- Direct form launching pattern (`frmExampleApp.Show`)
- Error handling with debug output
- Module-to-form communication

### code-behind.vba
Shows proper event handling:
- Form initialization
- Button click handlers  
- Form closure patterns

## 💡 Usage Tips

- Copy this structure for new applications
- Follow the naming conventions (btn, lbl, txt prefixes)
- Use direct form references after building
- Check Immediate window for debug output

For detailed documentation, see the main README.md in the repository root.