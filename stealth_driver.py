"""
stealth_driver.py  —  Shared anti-bot utilities for all GST / Income-Tax tools.

Import in any file:
    import sys, os
    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), <levels up to root>)))
    from stealth_driver import create_chrome_driver, show_browser_alert, STEALTH_CHROME_OPTIONS

EXE-safe: works when packaged with PyInstaller because Selenium 4.6+ ships
selenium-manager as a bundled binary (no internet download required at run time).
"""
import os
import sys

# ── Cache path: survives across runs even inside a frozen EXE ─────────────────
os.environ.setdefault("WDM_CACHE_PATH", os.path.join(os.path.expanduser("~"), ".wdm"))
os.environ.setdefault("WDM_LOCAL", "1")

try:
    from webdriver_manager.chrome import ChromeDriverManager as _CDM
    _HAS_WDM = True
except ImportError:
    _HAS_WDM = False

from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# ── Comprehensive stealth JS ───────────────────────────────────────────────────
# Injected via addScriptToEvaluateOnNewDocument so it runs BEFORE any page JS.
STEALTH_JS = """
    // 1. Hide automation flag
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});

    // 2. Full chrome runtime (portals check this)
    window.navigator.chrome = {
        runtime:    {},
        loadTimes:  function() {},
        csi:        function() {},
        app:        {}
    };

    // 3. Realistic plugin list
    Object.defineProperty(navigator, 'plugins', {
        get: () => [
            { name:'Chrome PDF Plugin',  filename:'internal-pdf-viewer',              length:1 },
            { name:'Chrome PDF Viewer',  filename:'mhjfbmdgcfjbbpaeojofohoefgiehjai', length:1 },
            { name:'Native Client',      filename:'internal-nacl-plugin',             length:2 }
        ]
    });

    // 4. Locale / platform
    Object.defineProperty(navigator, 'languages',           {get: () => ['en-IN','en','en-US']});
    Object.defineProperty(navigator, 'platform',            {get: () => 'Win32'});
    Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
    Object.defineProperty(navigator, 'deviceMemory',        {get: () => 8});

    // 5. Permissions API — headless Chrome reports 'denied' for notifications
    const _origQuery = window.navigator.permissions.query.bind(navigator.permissions);
    window.navigator.permissions.query = (p) =>
        p.name === 'notifications'
            ? Promise.resolve({state: Notification.permission})
            : _origQuery(p);
"""


def build_chrome_options(download_path=None):
    """Return a fully-configured ChromeOptions object with all anti-bot flags."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # Suppress Chrome console noise (CSP warnings, Angular errors, etc.)
    # so they don't appear in the tool window.
    options.add_argument("--log-level=3")
    options.add_argument("--silent")
    options.add_argument("--disable-logging")
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    prefs = {
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    if download_path:
        prefs["download.default_directory"] = str(download_path)
    options.add_experimental_option("prefs", prefs)
    return options


def create_chrome_driver(options=None):
    """
    EXE-safe Chrome driver factory.
    Priority:
      1. Selenium 4.6+ built-in selenium-manager  (works inside frozen EXE — no download).
      2. webdriver_manager fallback               (classic script mode).
    Injects full stealth JS and maximises the window automatically.
    """
    if options is None:
        options = build_chrome_options()

    driver = None
    # Try Selenium's bundled selenium-manager first
    try:
        driver = webdriver.Chrome(options=options)
    except Exception:
        pass

    # Fallback: webdriver_manager
    if driver is None and _HAS_WDM:
        try:
            driver = webdriver.Chrome(
                service=Service(_CDM().install()), options=options
            )
        except Exception:
            pass

    if driver is None:
        raise RuntimeError(
            "Could not start Chrome. "
            "Make sure Chrome is installed and Selenium >= 4.6 or webdriver-manager is available."
        )

    driver.maximize_window()
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": STEALTH_JS})
    return driver


def show_browser_alert(driver, message):
    """
    Inject a bright red banner INTO the browser page.
    Errors appear ON the website — never as a Python messagebox popup.
    Auto-dismisses after 10 seconds.
    Safe to call even if driver is None / already closed.
    """
    if not driver:
        return
    try:
        safe_msg = str(message).replace("'", "\\'").replace("`", "\\`").replace("\n", " ")
        js = f"""
        (function() {{
            var _id = '__gst_tool_alert__';
            var old = document.getElementById(_id);
            if (old) old.remove();
            var d = document.createElement('div');
            d.id = _id;
            d.innerText = '{safe_msg}';
            d.style.cssText = [
                'position:fixed','top:0','left:0','right:0',
                'z-index:2147483647',
                'background:#c0392b',
                'color:#fff',
                'font-size:15px',
                'font-weight:bold',
                'font-family:Arial,sans-serif',
                'padding:14px 24px',
                'text-align:center',
                'box-shadow:0 3px 12px rgba(0,0,0,0.5)',
                'letter-spacing:0.3px',
                'cursor:pointer'
            ].join(';');
            d.onclick = function() {{ d.remove(); }};
            document.body.prepend(d);
            setTimeout(function() {{ if (d.parentNode) d.remove(); }}, 10000);
        }})();
        """
        driver.execute_script(js)
    except Exception:
        pass  # Never crash the automation thread over a UI alert
