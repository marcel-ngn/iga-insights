# IGA-Insights

## Identity Governance & Administration Insights for SSO applications

IGA-Insights is a PowerShell script designed to help Identity & Access Management professionals optimize their Azure/Entra ID environment by providing actionable insights on SSO application usage, user access patterns, and license utilization.

![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![PowerShell: 7.0+](https://img.shields.io/badge/PowerShell-7.0+-blue.svg)
![Platform: Windows/Linux/macOS](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

## 🚀 Features

- **SSO Application Usage Analysis**: Identify active and inactive users for specific SSO applications
- **Comprehensive Reporting**: Generate detailed HTML dashboards and CSV exports
- **User Activity Tracking**: Track user sign-ins over customizable time periods
- **Group Membership Analysis**: Map group assignments to application access
- **Department & Job Title Insights**: Understand access patterns across your organization
- **Interactive Dashboards**: Filter and search through user activity data with ease
- **Visualizations**: See user activity, job title, and department distributions

## 📋 Prerequisites

- PowerShell 7.0 or higher
- Az PowerShell module (`Az.Accounts`, `Az.OperationalInsights`)
- Microsoft.Graph PowerShell module
- A Log Analytics Workspace with at least 365 days of retention period (recommended)
- SignInLogs data collection enabled in the workspace
- Azure/Entra ID account with:
  - Log Analytics Workspace access
  - Log Analytics Reader permissions
  - Azure/Entra ID Reader permissions

## 📥 Installation

```powershell
# Clone the repository
git clone https://github.com/marcel-ngn/iga-insights.git

# Navigate to the project directory
cd iga-insights

# Install required modules if not already installed
Install-Module -Name Az.Accounts, Az.OperationalInsights, Microsoft.Graph -Scope CurrentUser
```

## 🔧 Usage

### Analyze SSO Application Usage

```powershell
# Run the script with required parameters
.\Analyze-SSOUsage.ps1 -AppName "Figma" -ThresholdDays 90
```

### Parameters

| Parameter | Description | Required |
|-----------|-------------|----------|
| `-AppName` | Display name of the SSO application (Service Principal) to analyze | Yes |
| `-ThresholdDays` | Number of days to consider for determining user inactivity | Yes |

### Authentication Options

The script supports two authentication methods:
1. **User Credentials** - Interactive sign-in with your Azure account
2. **App Registration** - Non-interactive authentication using a service principal with certificate

## 📊 Report Outputs

The script generates two types of reports:

1. **HTML Dashboard** - An interactive web page with:
   - Summary statistics on active vs. inactive users
   - Charts showing activity distribution
   - Job title and department distribution visualizations
   - Searchable and filterable user details table
   - Pagination controls for large datasets

2. **CSV Export** - A detailed spreadsheet containing:
   - User principal names
   - Latest sign-in dates
   - Group memberships
   - Job titles and departments
   - Activity status

Reports are saved to your desktop by default, or to a custom location you specify.

## 📷 Screenshots

### Running the script in a Terminal
![image](https://github.com/user-attachments/assets/96d40870-ad63-42ac-9d2a-214f51981e77)
### Dashboard Overview
![image](https://github.com/user-attachments/assets/a636cb1c-1087-421b-809a-537610762bf6)


## 🔄 Workflow

1. Connect to Azure and Microsoft Graph services
2. Identify assigned groups for the specified application
3. Collect user information including group memberships
4. Query Log Analytics to determine active vs. inactive users
5. Generate comprehensive reports with visualizations
6. Present results in interactive HTML dashboard

## 🛠️ Advanced Configuration

### Using App Registration (Service Principal)

For automated or scheduled runs, you can use app registration authentication by:

1. Create an App Registration in Azure/Entra ID
2. Assign appropriate permissions (Log Analytics Reader, Directory.Read.All)
3. Generate a certificate for authentication
4. Update the script parameters or configuration file with your App details

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📜 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 📞 Contact

Marcel Nguyen - [@marcel_ngn](https://twitter.com/marcel_ngn)

Project Link: [https://github.com/marcel-ngn/iga-insights](https://github.com/marcel-ngn/iga-insights)

---
**Note**: This project is not affiliated with or endorsed by Microsoft.
