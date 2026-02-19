# SEMRUSH Data Extractor - Streamlit Web App

A modern web-based interface for extracting chart data from Semrush analytics pages.

## Installation & Setup

### Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Run the App

**Option A: Batch File (Windows)**
```bash
run_streamlit.bat
```

**Option B: PowerShell**
```powershell
.\.venv\Scripts\Activate.ps1
streamlit run streamlit_app.py
```

**Option C: Direct**
```bash
python -m streamlit run streamlit_app.py
```

App opens at: `http://localhost:8501`

## User Interface

### Layout Design
- **Left Column (Sidebar)**: "About This Platform" (always visible)
  - Platform overview
  - 7-step process guide
  - Quick reference

- **Right Column (Main)**: Workflow steps (stack sequentially)
  - Each step shows completion status  
  - Previous steps remain visible
  - Progressive disclosure of interface

### 7-Step Workflow

**Step 1: Enter Data Source URL**
- Input field with default Semrush URL
- "🚀 Run" button to initiate
- URL validation

**Step 2: Browser Opened**
- WebDriver initialization status
- Page loading progress
- Browser launch confirmation

**Step 3: Confirm Ready for Detection**
- Manual user confirmation required
- "✅ Yes, Detect Charts" button
- Security: prevents auto-detection

**Step 4: Detecting Charts**
- Real-time detection progress
- Chart count display when complete
- Status indicators

**Step 5: Select a Chart to Extract**
- Card-based chart display
- Shows title, width × height
- Individual "📊 Extract" buttons per chart

**Step 6: Extracting Data**
- Real-time progress bar
- Smart probing phase indication
- Detail scan phase status
- Data point counter

**Step 7: Download Your Data**
- Filename input field with default
- "📥 Download" button (direct download, no intermediate step)
- Auto-fit columns, professional formatting
- "📊 Extract Another Chart" or "🏠 Start Over" options

## Key Features

### Visual Design
- Bain & Company logo header on every page
- Professional color scheme (#dc143c crimson primary)
- Two-column responsive layout
- Step-by-step progress indication
- Color-coded status messages (green success, red errors, blue info)

### Smart Extraction Engine
- **Smart Probing**: Tests 11 positions × 7 Y-offsets = 77 test points
- **Intelligent Detail Scan**: Only scans detected data region ±8 positions
- **JavaScript Events**: Uses DOM PointerEvent/MouseEvent (no ActionChains)
- **Multi-Pattern Parsing**: Three fallback patterns for data extraction
- **Data Preservation**: Raw data sheet if parsing fails

### Session Management  
- **Driver Validation**: Checks if WebDriver session still active
- **Error Recovery**: Auto-recovery on session loss
- **Clean Shutdown**: Closes browsers on exit or crashes
- **State Persistence**: Maintains workflow state across interactions

### Excel Export
- **Structured Format**: Period × Metric pivot table
- **Fallback Preservation**: All 26+ tooltips captured even if parsing fails
- **Professional Styling**: Bold headers, borders, alignment
- **Custom Filenames**: User-defined file names
- **Direct Download**: Single click, no intermediate steps

## Troubleshooting & Support

**App won't start**
- Check Python version: `python --version` (need 3.7+)
- Clear Streamlit cache: `streamlit cache clear`

**Browser doesn't open**
- Navigate manually to: `http://localhost:8501`
- Ensure Chrome is installed

**Charts not detected**
- Wait for JavaScript to fully load
- Try waiting 10+ seconds before detecting
- Some pages require login

**Extraction unsuccessful**
- Click "Start Over" to reset session
- No data loss - raw data sheet preserved if parsing fails

**Excel download problems**
- Use simple filenames
- Clear browser cache

**Performance**
- Typical extraction: 3-5 minutes
- Close other apps for better performance

## Performance & Specifications

- **Extraction Time**: 3-5 minutes per chart
- **Data Capture**: Typically 20-50+ data points per chart
- **Success Rate**: 95%+ with smart probing and error handling
- **Memory Usage**: ~400-600MB (single Chrome browser instance)
- **Session Stability**: Auto-recovery on WebDriver loss

## Files

- **streamlit_app.py** - Main web application (845 lines)
  - Two-column layout with sidebar
  - 7-step guided workflow
  - Session state management
  - Error handling and recovery

- **bain_logo.png** - Header branding image
- **requirements.txt** - Dependencies
- **run_streamlit.bat** / **run_streamlit.ps1** - Quick launchers

## Platform Support

- **Browsers**: Chrome/Chromium (via Selenium)
- **OS**: Windows, macOS, Linux
- **Python**: 3.7+
- **Chrome**: Any recent version

## Tips for Success

1. **Default URL first** - Semrush URL is well-tested
2. **Patient page loading** - Give JavaScript time
3. **Modern charts** - SVG-based work best
4. **Clear tooltips** - Better data extraction
5. **Single extraction** - Each chart is fresh & independent

## Security & Privacy

- All processing local on your machine
- No data sent externally
- Anti-detection settings prevent site blocking
- Chrome manages updates automatically

---

**Last Updated:** February 20, 2026  
**Version:** 2.0 (Web App with Sidebar + Direct Download)
