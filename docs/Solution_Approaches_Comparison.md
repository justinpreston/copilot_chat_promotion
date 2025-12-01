# Copilot Chat Promotion Campaign: Solution Approaches

This document compares various implementation approaches for sending proactive Adaptive Card messages to Teams users as part of a Copilot Chat adoption campaign. Each approach balances complexity, reusability, cost, and IT admin friction differently.

---

## Executive Summary

| Approach | Complexity | Reusability | Cost | IT Admin Friction | Best For |
|----------|------------|-------------|------|-------------------|----------|
| **PowerShell Script (Current)** | Medium | Low | Free | High | One-time runs, technical admins |
| **Power Automate Flow** | Low-Medium | High | $15-150/mo | Medium | Orgs with Power Platform licenses |
| **Azure Logic Apps** | Medium | High | Pay-per-use (~$5-50/mo) | Medium | Azure-centric orgs, variable workloads |
| **Copilot Studio Agent + Flow** | Medium-High | Very High | $200/mo + Premium | Low | M365 Copilot licensed orgs |
| **Teams Toolkit Bot** | High | Very High | ~$5-50/mo Azure | Low (after deploy) | Large-scale, developer resources available |
| **Azure Automation Runbook** | Low | Medium | ~$0-5/mo | Low | Simple automation, Azure subscribers |

---

## Approach 1: PowerShell Script (Current Solution)

### How It Works
The current solution uses a PowerShell script that:
1. Authenticates to Microsoft Graph with delegated permissions
2. Reads user UPNs from an Excel file
3. Creates 1:1 chats with each user via Graph API
4. Sends Adaptive Card messages with Copilot prompts
5. Logs results to CSV

### Pros
- ✅ **No licensing cost** – Uses existing Graph permissions
- ✅ **Full control** – Can customize any aspect of the logic
- ✅ **Immediate** – Run directly from any machine with PowerShell 7+
- ✅ **Transparent** – All code is visible and auditable

### Cons
- ❌ **Requires PowerShell 7+** – Many enterprises still use Windows PowerShell 5.1
- ❌ **Module installation** – Needs `ImportExcel`, `Microsoft.Graph.Teams` modules
- ❌ **Interactive auth** – Requires user to log in; cannot run unattended
- ❌ **Hardcoded paths** – Must manually edit file paths for each environment
- ❌ **No scheduling** – Must be manually triggered each time
- ❌ **Manual Excel prep** – IT admin maintains user lists in Excel
- ❌ **Limited error visibility** – Logs to CSV, no dashboard

### Cost
**Free** (no licensing beyond existing M365)

### Best For
- Technical IT admins comfortable with PowerShell
- One-time or infrequent campaign sends
- Organizations with no Power Platform or Azure investment

---

## Approach 2: Power Automate Cloud Flow

### How It Works
A Power Automate flow replaces the PowerShell script with low-code automation:
1. **Trigger**: Manual button, scheduled recurrence, or SharePoint list item creation
2. **Data source**: SharePoint list or Excel Online table with user UPNs
3. **Graph API calls**: HTTP actions to create chats and send messages
4. **Logging**: Write results back to SharePoint list

### Architecture
```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  SharePoint     │────▶│  Power Automate  │────▶│  Microsoft      │
│  List (Users)   │     │  Cloud Flow      │     │  Graph API      │
└─────────────────┘     └──────────────────┘     └─────────────────┘
                               │                         │
                               ▼                         ▼
                        ┌──────────────────┐     ┌─────────────────┐
                        │  SharePoint      │     │  Teams 1:1      │
                        │  (Log Results)   │     │  Chat Messages  │
                        └──────────────────┘     └─────────────────┘
```

### Pros
- ✅ **Low-code** – Visual designer, no PowerShell knowledge needed
- ✅ **Scheduling built-in** – Run weekly campaigns automatically
- ✅ **SharePoint integration** – Easy user list management via browser
- ✅ **Error handling** – Built-in retry policies, run history
- ✅ **Shareable** – Export as solution package for other tenants
- ✅ **No infrastructure** – Runs in Microsoft cloud

### Cons
- ❌ **Premium license required** – HTTP connector needs Power Automate Premium ($15/user/month) or Process license ($150/month)
- ❌ **Action limits** – Standard license: ~40K actions/day; may need Process plan for large sends
- ❌ **App registration required** – Still need Azure AD app for Graph permissions
- ❌ **Throttling** – Graph API rate limits apply; need careful retry logic

### Cost
| Scenario | License | Monthly Cost |
|----------|---------|--------------|
| Single admin running campaigns | Per-user Premium | ~$15/month |
| High-volume (10K+ users/day) | Process plan | ~$150/month |
| Part of E5/Power Platform bundle | Included | $0 additional |

### Best For
- Organizations already invested in Power Platform
- Non-technical adoption/change management teams
- Recurring scheduled campaigns

---

## Approach 3: Azure Logic Apps

### How It Works
Identical workflow to Power Automate but runs as an Azure resource:
1. **Trigger**: Recurrence, HTTP webhook, or Azure Event Grid
2. **HTTP actions**: Same Graph API calls as Power Automate
3. **Logging**: Azure Storage, Cosmos DB, or SharePoint

### Architecture
```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Azure Storage  │────▶│  Azure Logic     │────▶│  Microsoft      │
│  (User List)    │     │  Apps            │     │  Graph API      │
└─────────────────┘     └──────────────────┘     └─────────────────┘
                               │                         │
                               ▼                         ▼
                        ┌──────────────────┐     ┌─────────────────┐
                        │  Log Analytics   │     │  Teams 1:1      │
                        │  (Monitoring)    │     │  Chat Messages  │
                        └──────────────────┘     └─────────────────┘
```

### Pros
- ✅ **Pay-per-use** – Only pay for actions executed (ideal for sporadic use)
- ✅ **No action limits** – Scale to any volume (cost scales linearly)
- ✅ **Azure ecosystem** – Integrates with Key Vault, Monitor, DevOps
- ✅ **Enterprise-grade** – RBAC, audit logs, compliance certifications
- ✅ **CI/CD support** – Deploy via ARM/Bicep templates, source control

### Cons
- ❌ **Azure subscription required** – Need Azure account and billing
- ❌ **Higher learning curve** – Requires Azure Portal familiarity
- ❌ **Cost can spike** – Very high volume = higher bills than fixed license
- ❌ **Less accessible** – Harder for business users to modify

### Cost
| Scenario | Actions/Month | Estimated Cost |
|----------|---------------|----------------|
| 500 users, weekly (4 sends) | ~10,000 | ~$1.25/month |
| 5,000 users, weekly | ~100,000 | ~$12.50/month |
| 50,000 users, weekly | ~1,000,000 | ~$125/month |

*Note: At very high volumes (~250K+ actions/day), Power Automate Process plan ($150/month) may be more economical.*

### Best For
- Azure-centric organizations
- Variable/unpredictable workloads
- DevOps teams wanting infrastructure-as-code

---

## Approach 4: Copilot Studio Agent + Power Automate/Logic Apps

### How It Works
A Copilot Studio agent provides a conversational interface for campaign management:
1. **Agent interface**: IT admin chats with agent in Teams: "Send Week 2 cards to licensed users"
2. **Agent actions**: Copilot Studio calls Power Automate flow via "Call an action"
3. **Flow execution**: Power Automate/Logic Apps handles the actual Graph API calls
4. **Status updates**: Agent reports back delivery results

### Architecture
```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  IT Admin       │────▶│  Copilot Studio  │────▶│  Power Automate │
│  (Teams Chat)   │     │  Agent           │     │  Cloud Flow     │
└─────────────────┘     └──────────────────┘     └─────────────────┘
        ▲                       │                         │
        │                       ▼                         ▼
        │               ┌──────────────────┐     ┌─────────────────┐
        └───────────────│  Status Updates  │     │  Graph API →    │
                        │  & Reports       │     │  Teams Messages │
                        └──────────────────┘     └─────────────────┘
```

### Pros
- ✅ **Natural language interface** – "Send the Researcher card to Sales team"
- ✅ **No technical knowledge** – Adoption teams can run campaigns via chat
- ✅ **Centralized orchestration** – Agent manages card selection, user targeting
- ✅ **Dynamic card generation** – Agent can customize cards per request
- ✅ **Highly reusable** – Export agent + flow as Power Platform solution

### Cons
- ❌ **Copilot Studio license required** – ~$200/month for 25K messages, or included with M365 Copilot
- ❌ **Still needs Premium connector** – Power Automate Premium for HTTP/Graph
- ❌ **Higher complexity** – Two systems to configure and maintain
- ❌ **M365 Copilot license for authors** – Admins managing the agent need Copilot licenses

### Cost
| Component | Cost |
|-----------|------|
| Copilot Studio | ~$200/month (or included with M365 Copilot) |
| Power Automate Premium | ~$15-150/month |
| **Total** | **~$215-350/month** |

### Best For
- Organizations with M365 Copilot licenses (Copilot Studio included)
- Teams wanting a "self-service" campaign tool for non-technical users
- Scenarios requiring dynamic card customization

---

## Approach 5: Teams Toolkit Proactive Notification Bot

### How It Works
A custom Teams bot built with Teams Toolkit sends proactive messages:
1. **Bot registration**: Created in Teams Developer Portal with notification-only capability
2. **Proactive installation**: Bot auto-installs for target users via Graph API
3. **Message sending**: Bot uses Bot Framework API to send Adaptive Cards
4. **Admin UI (optional)**: Teams Tab for campaign management

### Architecture
```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Admin UI       │────▶│  Azure Functions │────▶│  Bot Framework  │
│  (Teams Tab)    │     │  (Bot Logic)     │     │  Service        │
└─────────────────┘     └──────────────────┘     └─────────────────┘
                               │                         │
                               ▼                         ▼
                        ┌──────────────────┐     ┌─────────────────┐
                        │  Azure Storage   │     │  Teams 1:1      │
                        │  (State, Logs)   │     │  Bot Messages   │
                        └──────────────────┘     └─────────────────┘
```

### Pros
- ✅ **Teams-native** – Appears as a proper Teams app in the app catalog
- ✅ **Highest throughput** – Bot Framework API has minimal throttling
- ✅ **Auto-installation** – No manual user action to receive messages
- ✅ **Rich admin experience** – Can build full UI for campaign management
- ✅ **Serverless option** – Azure Functions consumption tier for low cost

### Cons
- ❌ **Development effort** – Requires JavaScript/TypeScript coding (~5-10 days)
- ❌ **Azure infrastructure** – Need Azure subscription, deployment pipeline
- ❌ **Ongoing maintenance** – Code updates, security patches
- ❌ **Initial admin consent** – Requires Global Admin to approve app permissions

### Cost
| Component | Cost |
|-----------|------|
| Azure Functions (Consumption) | ~$0-5/month |
| Azure Storage | ~$1-5/month |
| Azure Bot Service | Free tier |
| **Total** | **~$5-10/month** |

### Best For
- Organizations with developer resources
- Large-scale campaigns (10K+ users)
- Need for rich admin UI or complex scheduling
- Long-term, repeatable campaign infrastructure

---

## Approach 6: Azure Automation Runbook

### How It Works
Package the existing PowerShell script as an Azure Automation runbook:
1. **Runbook creation**: Upload PowerShell script to Azure Automation account
2. **Managed identity**: Use system-assigned identity for Graph API auth (no interactive login)
3. **Trigger**: Manual run, schedule, or webhook
4. **Logging**: Output to Azure Monitor logs

### Architecture
```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Azure Portal   │────▶│  Azure           │────▶│  Microsoft      │
│  (Start Job)    │     │  Automation      │     │  Graph API      │
└─────────────────┘     └──────────────────┘     └─────────────────┘
                               │                         │
                               ▼                         ▼
                        ┌──────────────────┐     ┌─────────────────┐
                        │  Azure Storage   │     │  Teams 1:1      │
                        │  (User List)     │     │  Chat Messages  │
                        └──────────────────┘     └─────────────────┘
```

### Pros
- ✅ **Minimal changes** – Reuse existing PowerShell script with minor edits
- ✅ **No interactive auth** – Managed identity handles authentication
- ✅ **Scheduling built-in** – Native schedule support
- ✅ **Very low cost** – 500 free minutes/month; pennies beyond that
- ✅ **Simple deployment** – Upload script, grant permissions, run

### Cons
- ❌ **Azure subscription required** – Need Azure account
- ❌ **PowerShell knowledge** – Still need to understand script for customization
- ❌ **Limited UI** – No rich campaign management interface
- ❌ **Module management** – Must import required modules to Automation account

### Cost
| Scenario | Cost |
|----------|------|
| <500 minutes/month | Free |
| 1,000 minutes/month | ~$2/month |
| 5,000 minutes/month | ~$10/month |

### Best For
- Quickest path to "unattended" execution
- Organizations with Azure subscription but no Power Platform
- IT admins who want to keep using PowerShell

---

## Decision Matrix

### By Organization Profile

| Organization Profile | Recommended Approach |
|---------------------|---------------------|
| No Azure, no Power Platform licenses | PowerShell Script (current) |
| Has Power Platform licenses, non-technical admins | Power Automate Flow |
| Azure-centric, DevOps culture | Logic Apps or Azure Automation |
| Has M365 Copilot licenses | Copilot Studio Agent |
| Developer resources, large scale | Teams Toolkit Bot |
| Minimal budget, Azure subscription | Azure Automation Runbook |

### By Use Case

| Use Case | Best Approach |
|----------|---------------|
| One-time pilot campaign | PowerShell Script |
| Weekly scheduled campaigns | Power Automate or Logic Apps |
| Self-service for adoption team | Copilot Studio Agent |
| Enterprise-wide rollout (50K+ users) | Teams Toolkit Bot |
| Quick automation with existing script | Azure Automation Runbook |

---

## Implementation Recommendations

### For Maximum Reusability (Sharing with Customers)
**Recommended: Power Automate Solution Package**

1. Create a Power Platform solution containing:
   - Cloud flow for sending messages
   - SharePoint list template for user management
   - Environment variables for configuration
   - Adaptive Card JSON files as attachments

2. Export as managed solution (.zip)

3. Customers import solution, configure:
   - Azure AD app registration (one-time)
   - SharePoint site for user lists
   - Environment variables (tenant ID, app ID)

**Estimated customer setup time: 30-60 minutes**

### For Lowest Barrier to Entry
**Recommended: Azure Automation Runbook + Documentation**

1. Create ARM template that deploys:
   - Automation account
   - Runbook with script
   - Managed identity with Graph permissions

2. Provide one-click "Deploy to Azure" button

3. Customer uploads user list to blob storage, clicks "Start"

**Estimated customer setup time: 15-30 minutes**

### For Maximum Flexibility
**Recommended: Hybrid Approach**

Provide multiple options:
1. **Quick start**: PowerShell script + documentation (current)
2. **Low-code**: Power Automate solution package
3. **Enterprise**: ARM template for Azure Automation or Logic Apps

Let customers choose based on their environment and skills.

---

## Next Steps

1. **Choose primary approach** based on target customer profile
2. **Build the solution** following the selected architecture
3. **Create deployment documentation** with step-by-step setup
4. **Test in pilot environment** before customer distribution
5. **Gather feedback** and iterate

---

## Appendix: Adaptive Card Assets

The following Adaptive Card templates are ready for use with any approach:

### Licensed Users (M365 Copilot)
| Week | Card | Focus |
|------|------|-------|
| 1 | `adaptiveCardActionsMicrosoft.json` | GPT-5 introduction, productivity prompts |
| 2 | `adaptiveCardActionsMicrosoft_Researcher.json` | Researcher agent deep dives |
| 3 | `adaptiveCardActionsMicrosoft_promptLikeACEO.json` | Executive-level prompting |
| 4 | `adaptiveCardActionsMicrosoft_PromptCoach.json` | Prompt Coach agent |

### Non-Licensed Users (Copilot Chat Web)
| Week | Card | Focus |
|------|------|-------|
| 1 | `adaptiveCardActionsMicrosoft_CopilotChat.json` | GPT-5 in Copilot Web |
| 2 | `adaptiveCardActionsMicrosoft_CopilotChat2.json` | Project plan review |
| 3 | `adaptiveCardActionsMicrosoft_multiImage.json` | Multi-image comparison |

---

*Document version: 1.0 | Last updated: December 2025*
