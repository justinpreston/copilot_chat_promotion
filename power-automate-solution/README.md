# Power Automate Solution for Copilot Chat Promotion

This folder contains a Power Automate cloud flow solution for sending promotional Adaptive Cards to Microsoft Teams users as part of a Copilot Chat adoption campaign.

## 🚀 Quick Start

1. **Download** `CopilotChatPromotion_1_0_0.zip` from this folder (or [GitHub Releases](https://github.com/luishdemetrio/copilot_chat_promotion/releases))

2. **Import** into Power Automate:
   - Go to [make.powerautomate.com](https://make.powerautomate.com)
   - Click **My flows** → **Import** → **Import Package (Legacy)**
   - Upload the zip file and click **Import**

3. **Create Azure AD App Registration** with Graph permissions (see [Deployment Guide](docs/Deployment-Guide.md))

4. **Run the flow** with your parameters and send cards to Teams users!

## Solution Overview

Instead of running PowerShell scripts manually, this solution provides:

- **Low-code automation** via Power Automate
- **7 pre-built Adaptive Cards** for Licensed and Non-Licensed users
- **OAuth authentication** to Microsoft Graph API
- **Importable package** for easy deployment

## Folder Structure

```text
power-automate-solution/
├── CopilotChatPromotion_1_0_0.zip     # 📦 Import this into Power Automate
├── README.md                          # This file
├── flows/
│   └── SendCopilotPromotionCards.json # Flow definition (reference)
├── sharepoint/
│   └── CopilotPromotionUsers-ListSchema.json # SharePoint list schema (optional)
├── config/
│   └── environment-variables.json     # Configuration reference
└── docs/
    └── Deployment-Guide.md            # Step-by-step setup instructions
```

## Flow Parameters

When you run the flow, you'll be prompted for:

| Parameter | Description | Example |
|-----------|-------------|---------|
| Campaign Week | Card week (1-4) | `1` |
| User Type | `Licensed` or `NonLicensed` | `Licensed` |
| Tenant ID | Your Azure AD tenant ID | `xxxxxxxx-xxxx-...` |
| Client ID | App registration client ID | `xxxxxxxx-xxxx-...` |
| Client Secret | App registration secret | `your-secret` |
| Target User UPN | Email of recipient | `user@contoso.com` |

## Requirements

| Requirement | Details |
|-------------|---------|
| License | Power Automate Premium (for HTTP connector) |
| Permissions | Application Admin (for app registration) |
| Azure AD | App registration with Graph API permissions |

## Campaign Cards

The flow includes 7 embedded Adaptive Cards:

### Licensed Users (M365 Copilot)

| Week | Card Theme |
|------|------------|
| 1 | GPT-5 Introduction |
| 2 | Researcher Agent |
| 3 | Prompt Like a CEO |
| 4 | Prompt Coach |

### Non-Licensed Users (Copilot Chat Web)

| Week | Card Theme |
|------|------------|
| 1 | GPT-5 in Copilot Web |
| 2 | Project Plan Review |
| 3 | Multi-Image Comparison |

## Architecture

```text
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Power Automate │────▶│  Microsoft       │────▶│  Teams 1:1      │
│  Manual Trigger │     │  Graph API       │     │  Chat Message   │
└─────────────────┘     └──────────────────┘     └─────────────────┘
        │                       │
        │                       ▼
        ▼               ┌─────────────────┐
 ┌──────────────┐       │  Adaptive Card  │
 │ User Inputs: │       │  Delivered to   │
 │ - Week       │       │  Target User    │
 │ - UserType   │       └─────────────────┘
 │ - Tenant ID  │
 │ - Client ID  │
 │ - Secret     │
 │ - User UPN   │
 └──────────────┘
```

## Flow Logic

1. **Trigger**: Manual button with input parameters
2. **Select Card**: Choose Adaptive Card based on UserType + Week
3. **Get Token**: OAuth to Microsoft Graph using app credentials
4. **Create Chat**: Create 1:1 chat with target user via Graph API
5. **Send Card**: Post Adaptive Card message to the chat
6. **Return Response**: Success status with recipient details

## Customization

### Adding New Cards

1. Export the flow from Power Automate
2. Edit the definition to add new switch cases in the `Select_Adaptive_Card` action
3. Re-import the modified flow

### Bulk Sending

The current flow sends to one user at a time. For bulk sending options:

- **Option A**: Run the flow multiple times with different UPNs
- **Option B**: Modify the flow to read from a SharePoint list (see `sharepoint/` folder for schema)
- **Option C**: Use the original PowerShell scripts in the `instructions/` folder

## Troubleshooting

See [docs/Deployment-Guide.md#troubleshooting](docs/Deployment-Guide.md#troubleshooting) for common issues and solutions.

## Contributing

Pull requests welcome! Please test thoroughly before submitting.

## License

MIT License - See repository root for details.
