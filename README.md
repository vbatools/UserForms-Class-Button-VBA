# clsButtonIcon - Double Icon Button Class for VBA UserForms

**English** | [Русский](README_RUS.md) | [UserForms-Class-ALL](https://github.com/vbatools/UserForms-Class-ALL/tree/main)

![Project Demo](User_Forms.gif)

## Description

`clsButtonIcon` is a VBA class that implements an interactive double icon button for use in Excel VBA UserForms. The class manages two Label controls (outer and inner), creating a unique visual element with the ability to change state and color on mouse hover.

## Features

- **Double icon**: Outer and inner icons, each with its own state
- **Two states**: Support for active and inactive states with different icons and colors
- **Hover effect**: Automatic highlighting on mouse hover
- **Flexible configuration**: Ability to customize colors, sizes, icons and behavior
- **Events**: Support for Click event to respond to user actions

## Requirements

- Microsoft Excel with VBA support
- "Segoe MDL2 Assets" font for icon display (included with Windows)

## Installation

1. Copy the `clsButtonIcon.cls` file to your VBA project
2. Import the class in the VBA editor

## Usage

### Basic Usage

```vba
Dim myButton As New clsButtonIcon

' Initialize button
Call myButton.Initialize(Label1, _
    ChrW$(59962), _        ' Outer icon (active state)
    ChrW$(59963), _        ' Outer icon (inactive state)
    RGB(0, 120, 215), _    ' Outer icon color (inactive state)
    RGB(255, 0, 0), _      ' Outer icon color (active state)
    ChrW$(59145), _        ' Inner icon (active state)
    ChrW$(60236), _        ' Inner icon (inactive state)
    RGB(128, 0, 128), _    ' Inner icon color (inactive state)
    RGB(128, 0, 128), _    ' Inner icon color (active state)
    True, _                ' Initial state
    True                   ' Enable hover effect
)
```

### Event Subscription

```vba
Private Sub myButton_Click(mLabelOut As MSForms.Label, mLabelIn As MSForms.Label, Value As Boolean)
    Debug.Print "Button state: " & Value
End Sub
```

### Changing Properties at Runtime

```vba
' Change icons
myButton.IconCodeOutOn = ChrW$(59145)
myButton.IconCodeOutOff = ChrW$(60236)

' Change colors
myButton.ColorOutOn = RGB(0, 176, 80)
myButton.ColorOutOff = RGB(128, 128, 128)

' Change state
myButton.Value = Not myButton.Value

' Control visibility and availability
myButton.Visible = True
myButton.Enabled = True
```

## Properties

| Property | Type | Description |
|----------|------|-------------|
| `IconCodeOutOn` | String | Outer element icon in active state |
| `IconCodeOutOff` | String | Outer element icon in inactive state |
| `IconCodeInOn` | String | Inner element icon in active state |
| `IconCodeInOff` | String | Inner element icon in inactive state |
| `ColorOutOn` | XlRgbColor | Outer icon color in active state |
| `ColorOutOff` | XlRgbColor | Outer icon color in inactive state |
| `ColorInOn` | XlRgbColor | Inner icon color in active state |
| `ColorInOff` | XlRgbColor | Inner icon color in inactive state |
| `SizeOut` | Single | Outer icon size |
| `SizeIn` | Single | Inner icon size |
| `Value` | Boolean | Logical button state |
| `Visible` | Boolean | Button visibility |
| `Enabled` | Boolean | Button availability |
| `HoverOn` | Boolean | Enable hover effect |

## Methods

| Method | Description |
|--------|-------------|
| `Initialize` | Initializes the button with specified parameters |
| `Version` | Returns class version information |

## Events

| Event | Parameters | Description |
|-------|------------|-------------|
| `Click` | `mLabelOut`, `mLabelIn`, `Value` | Occurs when button is clicked |

## Examples

Usage example is presented in the `frmTestClass.frm` form, which demonstrates all class capabilities.

To run the example:
1. Open the `button_v4.xlsm` file
2. Run the `showForm` macro from the `modShowForms.bas` module

## Error Handling

The class includes comprehensive error handling:

- **Initialization checks**: Validates that the provided Label control is not Nothing and has a parent form
- **Value validation**: Ensures sizes are greater than 0 and icon codes are not empty
- **Detailed error messages**: Provides specific error descriptions to help troubleshoot issues

## Recommendations for Improvement

1. **Add animation support between states**
   - Implement smooth transitions when changing button state
   - Use timers for gradual changes in transparency or color

2. **Improve size configuration flexibility**
   - Add property to configure the ratio between inner and outer icon sizes
   - Currently, inner icon size is fixed at 75% of the outer icon

3. **Add support for other icon fonts**
   - Extend functionality to support other popular icon fonts (e.g., FontAwesome, Material Icons)
   - Create enumeration for selecting icon font type

4. **Add bulk configuration methods**
   - Create `ConfigureTheme` method to configure all colors and sizes simultaneously
   - Add preset themes (e.g., light/dark)

5. **Improve performance with many buttons**
   - Optimize event handling when multiple class instances exist
   - Consider using a common event handler for multiple buttons

6. **Add state persistence support**
   - Implement `SaveState` and `LoadState` methods to save and restore current button state
   - Useful when closing and reopening forms

7. **Expand customization options**
   - Add properties to configure mouse hover effects
   - Provide option to configure effect delay time

8. **Add cloning methods**
   - Implement `Clone` method to create a copy of an existing button with the same parameters
   - This will simplify creating multiple identical buttons

9. **Improve documentation**
   - Add XML comments to all methods and properties
   - Create usage examples in documentation

## License

The project is distributed under the Apache License 2.0. See the LICENSE file for details.

## Author

VBATools