![GitHub License](https://img.shields.io/github/license/AhashSritharan/PCF-SmartButton?style=for-the-badge)
![GitHub Release](https://img.shields.io/github/v/release/AhashSritharan/PCF-SmartButton?style=for-the-badge)
![GitHub Downloads (all assets, all releases)](https://img.shields.io/github/downloads/AhashSritharan/PCF-SmartButton/total?style=for-the-badge)

# PCF Smart Button Control

A Power Apps Component Framework (PCF) control that provides dynamic, configurable buttons based on Dataverse settings. This control allows you to create flexible, context-aware buttons that can be used for custom actions and navigation in your Power Apps forms.

## Features

- **Dynamic Button Configuration**: Configure buttons directly in Dataverse without code changes
- **Visibility Conditions**: Show/hide buttons based on dynamic expressions
- **Custom Actions**: Execute custom scripts or navigate to URLs
- **Field Value Integration**: Use record field values in button labels, tooltips, and URLs
- **Smart Caching**: Optimized performance through intelligent caching of record data
- **Error Handling**: Robust error handling and graceful fallbacks
- **Fluent UI Integration**: Modern, consistent UI using Fluent UI components

## Installation

The easiest way to install the control is to download and import the solution from the [releases page](https://github.com/AhashSritharan/PCF-SmartButton/releases). You can choose either:
- **Managed Solution**: For production environments (recommended)
- **Unmanaged Solution**: For development environments

The solution package includes:
- The PCF Smart Button control
- Button Configuration table

## Configuration

### Button Configuration Table

Create records in the Button Configuration table with the following fields:

- **Name**: Unique identifier for the button configuration
- **Table Name**: Entity/table the button appears on (e.g., 'contact')
- **Button Label**: The text to display on the button
- **Button Tooltip**: Hover text for the button (optional)
- **Button Position**: Numeric order for button placement
- **Button Icon**: Fluent UI icon name (optional) - [Browse available icons](https://uifabricicons.azurewebsites.net/)
- **Visibility Expression**: Condition for showing/hiding button
- **URL**: The URL to navigate to when clicked (optional)
- **Show as**: Choose between Button or Link display
- **Action Script**: Custom script to execute on click (optional)

### Configuration Example

Here's an example of a fully configured button:

```
Name: Test Button 1
Table Name: contact
Button Label: This is a label {parentcustomerid.parentaccountid.websiteurl}
Button Tooltip: This is a tooltip {parentcustomerid.parentaccountid.websiteurl}
Button Position: 1
Button Icon: Send
Visibility Expression: {parentcustomerid.parentaccountid.websiteurl} != null
URL: {parentcustomerid.parentaccountid.websiteurl}
Show as: Button
Action Script: window.open("{parentcustomerid.parentaccountid.websiteurl}")
```

### Dynamic Values and Tokens

All text fields (Label, Tooltip, URL, Visibility Expression, and Action Script) support dynamic tokens:
- Use curly braces to reference field values: `{fieldname}`
- Access related records using dot notation: `{lookup.fieldname}`
- Chain multiple lookups: `{parentcustomerid.parentaccountid.websiteurl}`

### Advanced Features

#### Action Scripts
- Support for async/await operations
- Full access to form context and Xrm object
- Example with async operation:
  ```javascript
  await Xrm.Navigation.openUrl("{websiteurl}");
  ```

#### Button Icons
- Uses Fluent UI icons (formerly Fabric icons)
- Browse available icons at [UI Fabric Icons](https://uifabricicons.azurewebsites.net/)
- Use the friendly name from the icon gallery in the Button Icon field

#### Visibility Expressions
JavaScript expressions that determine button visibility:
```javascript
// Show only if website URL exists
{parentcustomerid.parentaccountid.websiteurl} != null

// Show based on status
{statuscode} === 1

// Complex conditions
{revenue} > 10000 && {statuscode} === 1
```

## Development

### Prerequisites

- Node.js
- npm
- Power Platform CLI (PAC)
- Visual Studio Code (recommended)

### Environment Setup

1. Install Power Platform CLI globally:
   ```bash
   npm install -g @microsoft/power-apps-cli
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create authentication profile:
   ```bash
   pac auth create --url https://yourorg.crm.dynamics.com
   ```

### Local Development

1. Create a new PCF project (if starting from scratch):
   ```bash
   pac pcf init --namespace YourNamespace --name YourControlName --template field
   ```

2. Build the control:
   ```bash
   npm run build
   ```

3. Start the test harness:
   ```bash
   npm start watch
   ```

### Deployment

1. Build the solution:
   ```bash
   pac solution build
   ```

2. Package the solution:
   ```bash
   pac solution pack
   ```

3. Push to your environment:
   ```bash
   pac pcf push --publisher-prefix YourPrefix
   ```

### Building

```bash
npm run build
```

### Testing

```bash
npm test
```

## Contributing

We welcome contributions! Please see our [contribution guide](CONTRIBUTING.md) for details.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For issues and feature requests, please [open an issue](https://github.com/AhashSritharan/PCF-SmartButton/issues) on GitHub.