# Sentinel Audit Dashboard V1.6 - New Features

This document describes the new features added in version 1.6 of the Sentinel Audit Dashboard.

## 🆕 New Features

### 1. Technician-Site Assignment System

Manage technicians and assign them to specific sites.

**Features:**
- Add, edit, and delete technicians
- Assign technicians to specific sites
- Store technician information including:
  - Name
  - Email address
  - Site assignment
  - Employee ID

**How to Use:**
1. Click the "👥 Technicians" button in the main dashboard
2. Use the "Add Technician" button to create new technician records
3. Assign each technician to a specific site from the dropdown
4. Edit or delete technicians as needed

### 2. AI Query Agent

Natural language interface to query dashboard data using Anthropic's Claude AI.

**Features:**
- Ask questions in plain English about your sensor data
- Get intelligent answers based on real-time database queries
- Examples of questions you can ask:
  - "How many failures are there on site K3?"
  - "What's the failure rate for Saffy this month?"
  - "Show me equipment with the most failures"
  - "How many tests were done on site EPL3 today?"

**Setup:**
1. Obtain an API key from [Anthropic](https://console.anthropic.com/)
2. Click "✉️ Email Config" button
3. Enter your Anthropic API key in the "AI Configuration" section
4. Save the configuration

**How to Use:**
1. Click the "🤖 AI Chat" button in the main dashboard
2. Type your question in the chat interface
3. Press Enter or click "Send"
4. Wait for the AI to analyze your data and respond

### 3. Automated Daily Email Reports

Automatically send daily failure reports to technicians at a scheduled time.

**Features:**
- Sends personalized daily reports to each technician
- Reports include:
  - Total tests for the day
  - Total failures for the day
  - Failure rate percentage
  - Detailed list of failed instruments with:
    - Equipment ID
    - Serial number
    - Equipment type
    - Test time
    - Gas type and measured value
- Scheduled delivery at configurable time (default: 5:00 AM)
- Separate reports for each site

**Setup:**

1. **Configure Email Settings:**
   - Click "✉️ Email Config" button
   - Enter your SMTP server details:
     - SMTP Server (e.g., smtp.gmail.com)
     - SMTP Port (usually 587 for TLS)
     - Your email address
     - Your email password (use app-specific password for Gmail)
   - Enable "Use TLS" (recommended)
   - Click "Test Email" to verify configuration

2. **Enable Daily Reports:**
   - In the Email Config window, check "Enable Daily Email Reports"
   - Set the desired send time in HH:MM format (e.g., "05:00")
   - Click "Save"

3. **Assign Technicians:**
   - Make sure technicians are assigned to sites (see Feature #1)
   - Each technician will receive reports only for their assigned site

**Gmail Users:**
For Gmail, you need to use an "App Password" instead of your regular password:
1. Go to your Google Account settings
2. Enable 2-Step Verification
3. Generate an App Password for "Mail"
4. Use this app password in the Email Config

## 📋 Installation

1. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Copy the example configuration file:
   ```bash
   cp config.json.example config.json
   ```

3. Edit `config.json` with your API keys and email settings

4. Run the application:
   ```bash
   python "Sentinel Audit Dashboard V1.6.py"
   ```

## 🔒 Security Notes

- The `config.json` file contains sensitive information (API keys, passwords)
- Keep this file secure and never commit it to version control
- For Gmail, always use App Passwords instead of your account password
- API keys and passwords are stored locally on your machine only

## 🐛 Troubleshooting

**AI Chat not working:**
- Verify your Anthropic API key is correct
- Check that the `anthropic` package is installed (`pip install anthropic`)
- Check the log file (`sentinel_audit_log.txt`) for errors

**Email not sending:**
- Click "Test Email" in Email Config to diagnose issues
- For Gmail: Ensure 2FA is enabled and you're using an App Password
- Check your SMTP settings and port number
- Check firewall settings if using corporate network
- Check the log file for detailed error messages

**Scheduler not running:**
- Verify "Enable Daily Email Reports" is checked in Email Config
- Check that the send time is in HH:MM format (e.g., "05:00")
- The application must be running for scheduled emails to be sent
- Consider running the application as a service for 24/7 operation

## 📝 Database Changes

V1.6 adds a new `technicians` table to each site database with the following schema:

```sql
CREATE TABLE IF NOT EXISTS technicians (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT NOT NULL,
    site_name TEXT NOT NULL,
    employee_id TEXT,
    created_date DATETIME DEFAULT CURRENT_TIMESTAMP
);
```

This table is automatically created when you first run V1.6 on an existing database.

## 🎯 Best Practices

1. **Technician Management:**
   - Keep technician email addresses up to date
   - Assign technicians to their primary sites
   - Use employee IDs to match with existing records

2. **AI Queries:**
   - Be specific in your questions (mention site names, date ranges)
   - Start with simple queries to understand capabilities
   - The AI has access to real-time data from all databases

3. **Email Reports:**
   - Test email configuration before enabling daily reports
   - Schedule reports for a time when technicians check email
   - Monitor the log file to ensure reports are being sent

## 💡 Future Enhancements

Potential future features:
- Multi-technician site assignments
- Custom report templates
- SMS notifications
- Historical trend analysis via AI
- Mobile app integration

## 📞 Support

For issues or questions:
1. Check the application log file: `sentinel_audit_log.txt`
2. Review this documentation
3. Contact the development team

---

**Version:** 1.6
**Release Date:** 2025-11-12
**Compatibility:** Python 3.8+
