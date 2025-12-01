# Power Automate Solution for Copilot Chat Promotion

This folder contains a Power Automate cloud flow solution for sending promotional Adaptive Cards to Microsoft Teams users as part of a Copilot Chat adoption campaign.

## Solution Overview

Instead of running PowerShell scripts manually, this solution provides:
- **Low-code automation** via Power Automate
- **SharePoint-based user management** (no Excel files)
- **Built-in scheduling and retry logic**
- **Exportable solution package** for easy customer deployment

## Folder Structure

```
power-automate-solution/
├── README.md                          # This file
├── flows/
│   └── SendCopilotPromotionCards.json # Main flow definition
├── sharepoint/
│   └── CopilotPromotionUsers-ListSchema.json # SharePoint list schema
├── config/
│   └── environment-variables.json     # Configuration variables
├── solution/
│   └── solution.xml                   # Power Platform solution manifest
└── docs/
    └── Deployment-Guide.md            # Step-by-step setup instructions
```

## Quick Start

1. **Read the deployment guide**: [docs/Deployment-Guide.md](docs/Deployment-Guide.md)

2. **Create Azure AD App Registration** with Graph permissions:
   - `Chat.Create`
   - `Chat.ReadWrite.All`
   - `User.Read.All`

3. **Create SharePoint list** using schema in `sharepoint/`

4. **Import solution** to Power Automate and configure environment variables

5. **Add users** to SharePoint list and run the flow

## Requirements

| Requirement | Details |
|-------------|---------|
| License | Power Automate Premium ($15/user/month) |
| Permissions | Global Admin or Application Admin (for app registration) |
| SharePoint | Site with list creation permissions |
| Azure AD | Ability to create app registrations |

## Campaign Cards

The flow includes embedded Adaptive Cards for:

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

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  SharePoint     │────▶│  Power Automate  │────▶│  Microsoft      │
│  Users List     │     │  Cloud Flow      │     │  Graph API      │
└─────────────────┘     └──────────────────┘     └─────────────────┘
                               │                         │
                               │                         ▼
                               │                 ┌─────────────────┐
                               │                 │  Teams 1:1      │
                               │                 │  Chat Messages  │
                               │                 └─────────────────┘
                               ▼
                        ┌──────────────────┐
                        │  SharePoint      │
                        │  Status Updates  │
                        └──────────────────┘
```

## Flow Logic

1. **Trigger**: Manual button with parameters (Week, UserType, TestMode)
2. **Get Token**: OAuth to Microsoft Graph using app credentials
3. **Select Card**: Choose Adaptive Card based on Week + UserType
4. **Get Users**: Query SharePoint for users with Status = "Pending"
5. **For Each User**:
   - Create 1:1 chat via Graph API
   - Send Adaptive Card message
   - Update SharePoint with result (Sent/Failed)
6. **Return Summary**: Success and failure counts

## Customization

### Adding New Cards

Edit `flows/SendCopilotPromotionCards.json` to add new switch cases:

```json
"NewCard_Week5": {
  "case": "Licensed_Week5",
  "actions": {
    "Set_Card_Licensed_Week5": {
      "type": "SetVariable",
      "inputs": {
        "name": "AdaptiveCardJson",
        "value": "{ your card JSON here }"
      }
    }
  }
}
```

### Changing SharePoint Columns

Update `sharepoint/CopilotPromotionUsers-ListSchema.json` and recreate the list.

## Troubleshooting

See [docs/Deployment-Guide.md#troubleshooting](docs/Deployment-Guide.md#troubleshooting) for common issues and solutions.

## Contributing

Pull requests welcome! Please test thoroughly before submitting.

## License

MIT License - See repository root for details.
