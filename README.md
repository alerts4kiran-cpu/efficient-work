# Amazon Connect - Combined Monitoring Script

A comprehensive Tampermonkey script for Amazon Connect that provides real-time agent highlighting, activity tracking, and CSV export functionality with optimized performance.

## 🚀 Features

### Agent Highlighting
Color-coded highlighting based on availability and break duration:
- 🟡 **Yellow**: Available 2-5 minutes
- 🔴 **Red**: Available 5+ minutes
- 🔵 **Blue**: Break exceeding threshold (default: 20 minutes)
- 🟠 **Orange**: Lunch exceeding threshold (default: 30 minutes)

### Real-Time Activity Summary
- **Live counter** showing all agent activities
- Updates every second automatically
- **Draggable box** - reposition anywhere on screen
- **Download button** - Export activity summary as CSV
- Shows all activities including Available, Break, Lunch, Personal, Training, Meeting, etc.

### CSV Export Options
- **📥 Download All** - Export all agent data with activity summary
- **⭐ Download Highlighted** - Export only highlighted agents
- **Smart filtering** - Excludes duplicate activity data from agent records
- **Comprehensive summary** - Properly formatted activity summary at end of file

### Customizable Settings
- Adjust highlight thresholds for each color
- Settings persist across sessions
- Easy-to-use settings dialog

### Optimized Performance
- Minimal browser impact
- Efficient caching and throttling
- Works smoothly even with large agent lists

---

## 📦 Installation

### Prerequisites
Install **Tampermonkey** extension for your browser:

| Browser | Installation Link |
|---------|------------------|
| 🦊 **Firefox** | [Install Tampermonkey](https://addons.mozilla.org/en-US/firefox/addon/tampermonkey/) |
| 🌐 **Chrome** | [Install Tampermonkey](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo) |
| 🔷 **Edge** | [Install Tampermonkey](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd) |

 🦊 **Firefox** --- Instalation is simple. Add the extension to the firefox. Install the script. Make sure Tamper monkey is enabled and the script "Connect - Combined Highlight, Download & Activity Counter (Optimized)" is turned on. You will see a toggle button next to script. Toggle it to green to run the script.
 
 🌐 **Chrome** --- Install the tamper monkey extension. Once installed, Click on Extensions icon and then Manage Extension. Find Tampermonkey and click on Details option under tampermonkey. Once done, make sure that tampermonkey is turned on. You will see the toggle switch on the right side. 
 In the menu, 
 1. You will see drop down menu for "Allow this extension to read and change all your data on websites you visit" . Select "On all sites"
 2. Make sure "Allow User Scripts" is toggled on
 3. "Allow Access to files URLs" is toggled on
 4. "Allow in incognito" is toggeled on (Optional)
Once done, refresh the connect page. 

🔷 **Edge** --- Install the tamper monkey extension. Once installed, Click on Extensions icon and then Manage Extension. Find Tampermonkey and click on Details option under tampermonkey. Once done, make sure that tampermonkey is turned on. You will see the toggle switch on the right side. 
 In the menu, 
 1. In the drop down menu of "Site access" - Select "On all sites"
 2. Check box Allow in Private (Optional)
 3. Check box "Allow access to file URLs"
Once done, refresh the connect page.

### One-Click Installation

**[📥 Click here to install the script](https://raw.githubusercontent.com/alerts4kiran-cpu/efficient-work/main/Connect%20-%20Combined%20Highlight%2C%20Download%20%26%20Activity%20Counter%20(Optimized)-4.9.user.js)**

Tampermonkey will automatically detect the script and prompt you to install it.

### Manual Installation Steps

1. Click the installation link above
2. Tampermonkey will open with the script preview
3. Click the **"Install"** button
4. Navigate to your Amazon Connect Real-Time Metrics page
5. The script will activate automatically

After installation, Make sure to that the tampermonkey extension is enabled and Connect - Combined Highlight, Download & Activity Counter (Optimized) is turned on (Toggle button would be green). Refresh the connect page once and you should see the script running.
---

## 🎯 Usage

### Activity Summary Box
1. Click **📊 Activity** button (top right) to toggle the activity summary
2. **Drag** the box to reposition it anywhere on screen
3. Click **⬇** button inside the box to download activity summary CSV
4. Box updates automatically every second

### Download Options
- **📥 Download All** - Exports all agent data with activity summary
- **⭐ Download Highlighted** - Exports only color-highlighted agents
- Both options include properly formatted activity summary at the end

### Customize Settings
1. Click **⚙️ Highlight Settings** button (top right)
2. Adjust duration thresholds for each highlight color:
   - Yellow (Available): Min and Max duration
   - Red (Available): Minimum duration
   - Blue (Break): Minimum duration
   - Orange (Lunch): Minimum duration
3. Click **Save** - settings are stored automatically

---

## 🔄 Automatic Updates

The script checks for updates automatically through Tampermonkey.

### Enable Auto-Updates
1. Open **Tampermonkey Dashboard** (click extension icon → Dashboard)
2. Find the script in the list
3. Click the script name
4. Go to **"Settings"** tab
5. Set **"Check for updates"** to **"On"**
6. Set **"Update interval"** to **"Daily"** (recommended)

### Manual Update Check
- In Tampermonkey dashboard, click the **"Last updated"** timestamp
- Or click **"Check for updates"** button

When an update is available, Tampermonkey will notify you automatically.

---

## 🖥️ Compatibility

| Browser | Status |
|---------|--------|
| Firefox | ✅ Fully supported |
| Chrome | ✅ Fully supported |
| Microsoft Edge | ✅ Fully supported |

**Requirements:**
- Tampermonkey extension installed
- Access to Amazon Connect Real-Time Metrics page

---

## 📊 CSV Export Details

### Agent Data Section
Includes all standard columns:
- Agent Login, Channels, Activity, Duration
- Agent Hierarchy, Routing Profile
- Contact State, Queue, Occupancy
- And more...

### Activity Summary Section
Properly formatted summary at the end of the file:
