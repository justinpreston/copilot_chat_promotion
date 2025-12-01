# Power Automate Deployment Guide
## Copilot Chat Promotion Campaign

This guide walks you through deploying the Power Automate solution for sending Copilot Chat promotional Adaptive Cards to Teams users.

---

## Prerequisites

Before you begin, ensure you have:

- [ ] **Global Administrator** or **Application Administrator** role in Microsoft Entra ID (for app registration)
- [ ] **Power Automate Premium** license (required for HTTP connector)
- [ ] **SharePoint** access to create a site and list
- [ ] **Teams** access to verify message delivery

---

## Deployment Steps

### Step 1: Create Azure AD App Registration

The flow needs application permissions to send messages via Microsoft Graph API.

1. Go to [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations**

2. Click **New registration**
   - **Name**: `Copilot Chat Promotion Flow`
   - **Supported account types**: `Accounts in this organizational directory only`
   - Click **Register**

3. On the app overview page, copy these values:
   - **Application (client) ID** → You'll need this as `ClientId`
   - **Directory (tenant) ID** → You'll need this as `TenantId`

### Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**

2. Click **Add a permission** → **Microsoft Graph** → **Application permissions**

3. Add these permissions:
   | Permission | Purpose |
   |------------|---------|
   | `Chat.Create` | Create 1:1 chats with users |
   | `Chat.ReadWrite.All` | Send messages to chats |
   | `User.Read.All` | Validate user existence |

4. Click **Grant admin consent for [Your Tenant]**

5. Wait for green checkmarks to appear next to each permission

   ![API Permissions Screenshot](../images/api-permissions.png)

### Step 3: Create Client Secret

1. In your app registration, go to **Certificates & secrets**

2. Click **New client secret**
   - **Description**: `Copilot Promotion Flow`
   - **Expiration**: `24 months` (or per your security policy)

3. Click **Add**

4. ⚠️ **IMMEDIATELY copy the Value** → This is your `ClientSecret`
   - You will NOT be able to see this value again!
   - Store it securely (e.g., Azure Key Vault, password manager)

### Step 4: Create SharePoint Site and List

1. Go to [SharePoint](https://www.office.com/launch/sharepoint)

2. Create a new **Communication site** or **Team site**
   - **Name**: `Copilot Promotion` (or your preferred name)
   - Copy the site URL (e.g., `https://contoso.sharepoint.com/sites/CopilotPromotion`)

3. Create a new list:
   - Click **New** → **List** → **Blank list**
   - **Name**: `CopilotPromotionUsers`

4. Add these columns:

   | Column Name | Type | Required | Notes |
   |-------------|------|----------|-------|
   | Title | Single line of text | No | Display name (optional) |
   | UPN | Single line of text | Yes | User's email (e.g., john@contoso.com) |
   | UserType | Choice | Yes | Options: `Licensed`, `NonLicensed` |
   | Status | Choice | No | Options: `Pending`, `Sent`, `Failed`, `Skipped` |
   | WeekSent | Single line of text | No | Week number sent (1-4) |
   | SentDate | Date and time | No | When message was sent |
   | ErrorMessage | Single line of text | No | Error details if failed |
   | Department | Single line of text | No | For filtering/reporting |

5. Set default value for **Status** to `Pending`

### Step 5: Import the Power Automate Solution

1. Go to [Power Automate](https://make.powerautomate.com)

2. Click **Solutions** in the left navigation

3. Click **Import solution** → **Browse**

4. Select `CopilotChatPromotion_1_0_0.zip`

5. Click **Next**

6. Configure environment variables:

   | Variable | Value |
   |----------|-------|
   | Tenant ID | Your tenant ID from Step 1 |
   | App Client ID | Your client ID from Step 1 |
   | App Client Secret | Your secret from Step 3 |
   | SharePoint Site URL | Your site URL from Step 4 |

7. Configure connections:
   - **SharePoint**: Create new or select existing connection
   - Sign in with an account that has access to the SharePoint site

8. Click **Import**

### Step 6: Verify the Flow

1. In the imported solution, find **Send Copilot Promotion Cards** flow

2. Click to open, then click **Edit**

3. Verify all connections show green checkmarks

4. Click **Save** if you made any changes

---

## Running the Campaign

### Add Users to the List

1. Go to your SharePoint list (`CopilotPromotionUsers`)

2. Add users manually or import from CSV:
   - Click **⋮** → **Export to Excel** to get template
   - Fill in UPNs and UserType
   - Copy/paste back into SharePoint

3. Ensure **Status** is `Pending` for users to receive messages

### Trigger the Flow

1. Go to [Power Automate](https://make.powerautomate.com) → **My flows**

2. Find **Send Copilot Promotion Cards**

3. Click **Run** (play button)

4. Configure the run:

   | Parameter | Description |
   |-----------|-------------|
   | Campaign Week | Which week's card to send (1-4) |
   | User Type | `Licensed` or `NonLicensed` |
   | Test Mode | If `true`, only sends to first 5 users |

5. Click **Run flow**

### Monitor Progress

1. Click on the running flow to see real-time progress

2. Check SharePoint list for updated statuses:
   - `Sent` = Message delivered successfully
   - `Failed` = Error occurred (check ErrorMessage column)

3. View run summary at the end with success/failure counts

---

## Campaign Schedule

For a typical 4-week campaign:

| Week | Licensed Users Card | Non-Licensed Users Card |
|------|---------------------|-------------------------|
| 1 | GPT-5 introduction | GPT-5 in Copilot Web |
| 2 | Researcher agent | Project plan review |
| 3 | Prompt Like a CEO | Multi-image comparison |
| 4 | Prompt Coach | (3 weeks only) |

**Recommended schedule**: Run once per week, same day/time for consistency.

---

## Troubleshooting

### Common Issues

#### "401 Unauthorized" when getting access token
- **Cause**: App registration permissions not granted
- **Fix**: Ensure admin consent was granted in Step 2

#### "403 Forbidden" when sending message
- **Cause**: Missing `Chat.ReadWrite.All` permission
- **Fix**: Add permission and grant admin consent

#### "User not found" errors
- **Cause**: UPN is incorrect or user doesn't exist
- **Fix**: Verify email addresses in SharePoint list

#### Flow runs but no messages received
- **Cause**: Status column not set to "Pending"
- **Fix**: Reset Status to "Pending" for target users

### Checking Logs

1. Go to **Power Automate** → **My flows** → **Send Copilot Promotion Cards**

2. Click **Run history** (clock icon)

3. Click on a specific run to see step-by-step details

4. Expand failed steps to see error messages

---

## Cost Estimate

| Component | Cost |
|-----------|------|
| Power Automate Premium | ~$15/user/month or $150/month (Process plan) |
| SharePoint | Included with M365 |
| Azure AD App | Free |
| **Total** | **$15-150/month** |

---

## Security Considerations

1. **Client Secret**: Store in Azure Key Vault for production
2. **Permissions**: Use least-privilege (only required Graph permissions)
3. **Access Control**: Limit who can run the flow (co-owners only)
4. **Audit**: Monitor Azure AD sign-in logs for app activity
5. **Rotation**: Rotate client secret before expiration

---

## Support

- **Issues**: [GitHub Issues](https://github.com/luishdemetrio/copilot_chat_promotion/issues)
- **Documentation**: [Full README](https://github.com/luishdemetrio/copilot_chat_promotion)

---

*Last updated: December 2025*
