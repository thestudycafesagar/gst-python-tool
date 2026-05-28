"""
GSTR-1 Downloader â€” Python Automation
=======================================
Direct HTTP against the GST portal (no GSP, no Selenium).
Mirrors GstOnlineActivity.GetGstr1WithOutOTP() from the C# codebase.

Two download strategies (tried in order):
  1. Offline ZIP download â€” `/offline/download/generate` + `/offline/download/url`
     Returns the complete filed return as a ZIP â†’ JSON file.
  2. Per-section API fallback â€” summary + per-section invoice calls.

Flow:
  1. get_captcha_base64()
  2. login(username, password, captcha)
     â””â”€ login_with_otp(otp)  if OTP_REQUIRED
  3. download_gstr1(period, year)
  4. logout()

Install: pip install requests Pillow
"""

import io
import json
import zipfile
import random
import base64
import logging
import time
from pathlib import Path
from typing import Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# Constants
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)

MFP_JSON = json.dumps({
    "VERSION": "2.1",
    "MFP": {
        "Browser": {
            "UserAgent": (
                "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; "
                "Trident/6.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; "
                ".NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)"
            ),
            "CookieEnabled": True
        },
        "IEPlugins": {
            "Flash": "20,0,0,306", "WindowsMediaPlayer": "12,0,7601,17514",
            "VBVersion": "10.0.16521", "ConnectionType": "lan",
            "AddressBook": "6,1,7601,17514", "BrowsingPack": "10,0,9200,16521",
            "DHTMLDataBinding": "10,0,9200,16521", "IEHelp": "10,0,9200,16521",
            "IEHelpEngine": "6,2,9200,16521", "OfflineBrowsingPack": "10,0,9200,16521",
            "OutlookExpress": "6,1,7601,17514", "WindowsDektopUpdate": "6,1,7601,17514"
        },
        "NetscapePlugins": {},
        "Screen": {
            "FullHeight": 768, "AvlHeight": 724, "FullWidth": 1366, "AvlWidth": 1366,
            "BufferDepth": 0, "ColorDepth": 24, "PixelDepth": 24,
            "DeviceXDPI": 96, "DeviceYDPI": 96, "FontSmoothing": True, "UpdateInterval": 0
        },
        "System": {"Platform": "Win32", "OSCPU": "x86", "userLanguage": "en-IN", "Timezone": -330}
    },
    "ExternalIP": "",
    "MESC": {"mesc": "mi=2;cd=150;id=30;mesc=207943;mesc=224290"}
}, separators=(',', ':'))

GST_ERRORS = {
    "SWEB_9000": "Invalid Captcha. Please try again.",
    "AUTH_9002": "Invalid UserId or Password.",
    "AUTH_9033": "Password has Expired. Please reset your password.",
    "SWEB_8000": "Error at GSTN site. Please try after sometime.",
    "SWEB_9014": "Account locked. Reset password first.",
    "RSK_1000":  "OTP Required",
    "SWEB_9003": "Wrong OTP entered.",
}

# Sections downloaded via per-invoice API when offline ZIP is unavailable
GSTR1_FLAT_SECTIONS = [
    ("B2CL",   "b2cl",   "SU"),
    ("B2CS",   "b2cs",   "SU"),
    ("EXP",    "exp",    "SU"),
    ("NIL",    "nil",    None),
    ("CDNR",   "cdnr",   "SU"),
    ("CDNUR",  "cdnur",  "SU"),
    ("HSN",    "hsnsum", "SU"),
    ("AT",     "at",     "SU"),
    ("TXP",    "txp",    "SU"),
    ("DOC_ISSUE", "dociss", "SU"),
    ("ECOM",   "ecom",   "SU"),
    ("CDNURA", "cdnura", "SU"),
]


# =============================================================================
# Result classes
# =============================================================================

class LoginResult:
    SUCCESS      = "success"
    OTP_REQUIRED = "otp_required"
    FAILED       = "failed"

    def __init__(self, status: str, message: str):
        self.status  = status
        self.message = message

    @property
    def is_success(self) -> bool:
        return self.status == self.SUCCESS

    @property
    def otp_required(self) -> bool:
        return self.status == self.OTP_REQUIRED

    def __repr__(self):
        return f"LoginResult(status={self.status!r}, message={self.message!r})"


class Gstr1Result:
    """
    Holds downloaded GSTR-1 data.

    If offline ZIP download succeeded â†’ zip_jsons has the raw JSON strings
    from inside the ZIP file(s), and source == "zip".

    If fallback per-section used â†’ sections dict has section data,
    and source == "api".
    """
    def __init__(self):
        self.source: str      = ""          # "zip" or "api"
        self.zip_jsons: list  = []          # raw JSON(s) from inside zip
        self.summary_json: str = ""         # /gstr1/summary response (always fetched)
        self.sections: dict   = {}          # section_name â†’ JSON string (api fallback)
        self.b2b_count: int   = 0           # number of B2B suppliers fetched (api mode)
        self.period: str      = ""
        self.year: int        = 0
        self.rtn_prd: str     = ""
        self.ustatus_json: str = ""         # /ustatus or profile response

    @property
    def main_json(self) -> str:
        """Best available data â€” ZIP JSON or summary JSON."""
        if self.zip_jsons:
            return self.zip_jsons[0]
        return self.summary_json

    def __repr__(self):
        return (f"Gstr1Result(period={self.period}, year={self.year}, "
                f"source={self.source!r}, zip_files={len(self.zip_jsons)})")


# =============================================================================
# Downloader
# =============================================================================

class Gstr1Downloader:
    """
    GSTR-1 downloader â€” direct portal HTTP, no GSP.

    Primary strategy: offline ZIP download (complete filed return data).
    Fallback: per-section invoice API calls + summary.
    """

    _URL_CAPTCHA      = "https://services.gst.gov.in/services/captcha"
    _URL_AUTHENTICATE = "https://services.gst.gov.in/services/authenticate"
    _URL_RBA_OTP      = "https://services.gst.gov.in/services/validate/rba/otp"
    _URL_USTATUS      = "https://services.gst.gov.in/services/api/ustatus"
    _URL_DROPDOWN     = "https://return.gst.gov.in/returns/auth/api/dropdown"
    _URL_LOGOUT       = "https://services.gst.gov.in/services/logout"
    _BASE_RETURN      = "https://return.gst.gov.in"

    def __init__(self, log_path: str = "", timeout: int = 180,
                 session: Optional[requests.Session] = None,
                 log_callback=None):
        self._timeout        = timeout
        self._log_path       = log_path
        self._device_id      = ""
        self._last_login_raw = ""
        self._profile_data   = {}
        self._log_callback   = log_callback  # callable(msg: str) forwarded to GUI
        self._yearly_stop_source = None
        self._yearly_stop_val = False

        if session is not None:
            self._session = session
            self._session_external = True
        else:
            self._session = requests.Session()
            self._session_external = False
            retry = Retry(total=5, backoff_factor=1,
                          status_forcelist=[500, 502, 503, 504],
                          allowed_methods=["GET", "POST"])
            self._session.mount("https://", HTTPAdapter(max_retries=retry))
            self._session.mount("http://",  HTTPAdapter(max_retries=retry))

        self._setup_logging()
        self._log("Gstr1Downloader ready.")

    @property
    def _yearly_stop(self) -> bool:
        if self._yearly_stop_source and hasattr(self._yearly_stop_source, "_yearly_stop"):
            return self._yearly_stop_source._yearly_stop
        return getattr(self, "_yearly_stop_val", False)

    @_yearly_stop.setter
    def _yearly_stop(self, val: bool):
        self._yearly_stop_val = val
    # â”€â”€ Step 2.5 â€” Profile â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def fetch_profile(self) -> dict:
        """Fetches registration profile (Legal Name, GSTIN, etc)."""
        self._log("Fetching user profile...")
        try:
            # Warm up navigation to profile page
            self._session.get("https://services.gst.gov.in/services/auth/myprofile", 
                              headers=self._common_headers(), verify=False, timeout=self._timeout)
            
            raw = self._post_json("https://services.gst.gov.in/services/auth/profile/detail", {}, 
                                  referer="https://services.gst.gov.in/services/auth/myprofile")
            data = json.loads(raw)
            if data and isinstance(data, dict):
                self._profile_data = data
                self._log(f"Profile fetched: {data.get('lgl_nm')}")
                return data
        except Exception as e:
            self._log(f"Profile fetch error: {e}")
        return {}

    # â”€â”€ Step 1 â€” Captcha â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def get_captcha_base64(self) -> str:
        if not self._session_external:
            self._session.cookies.clear()
        self._device_id = ""
        self._last_login_raw = ""

        url = f"{self._URL_CAPTCHA}?rnd={random.random()}"
        resp = self._session.get(url, headers={
            "Accept":           "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
            "Referer":          "https://services.gst.gov.in/services/login",
            "X-Requested-With": "XMLHttpRequest",
            "Sec-Fetch-Dest":   "image",
            "Sec-Fetch-Mode":   "no-cors",
            "Sec-Fetch-Site":   "same-origin",
            "Cache-Control":    "no-cache",
            "User-Agent":       USER_AGENT,
        }, verify=False, timeout=self._timeout)
        resp.raise_for_status()
        self._log(f"Captcha fetched ({len(resp.content)} bytes).")
        return base64.b64encode(resp.content).decode("utf-8")

    # â”€â”€ Step 2 â€” Login â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def login(self, username: str, password: str, captcha_text: str) -> LoginResult:
        if len(password) > 15:
            return LoginResult(LoginResult.FAILED,
                               "Password too long â€” GST portal max 15 characters.")
        body = {
            "username": username, "password": password,
            "captcha": captcha_text, "mFP": MFP_JSON,
            "deviceID": self._device_id if self._device_id else None,
            "type": "username"
        }
        self._log(f"Logging in as: {username}")
        raw = self._post_json(self._URL_AUTHENTICATE, body,
                              referer="https://services.gst.gov.in/services/login")
        self._last_login_raw = raw
        self._extract_device_id(raw)
        return self._parse_login_response(raw)

    def login_with_otp(self, otp: str) -> LoginResult:
        device_id = self._device_id
        try:
            prev = json.loads(self._last_login_raw)
            if prev.get("deviceID"):
                device_id = str(prev["deviceID"])
        except Exception:
            pass

        body = {
            "username": None, "password": None, "captcha": None,
            "mFP": MFP_JSON, "deviceID": device_id, "type": "username",
            "applIP": None, "email": None, "emailOtp": None,
            "mobileNo": None, "oidar": None, "otpAuthSts": "true",
            "otpDetails": json.dumps({"otp": otp}), "refNum": None,
            "riskLvl": "INCREASEAUTH", "role": None,
            "ruleAnnotation": (
                "UNKNOWNUSER=N;MFPMISMATCH=N;NEGATIVECOUNTRY=N;"
                "UNTRUSTEDIP=N;UNKNOWNDEVICEID=N;DEVICEVELOCITY=N;"
                "ZONEHOPPING=N;USERDEVICENOTASSOCIATED=Y;GST_USER_VELOCITY=N;"
            ),
            "ruleMatched": "USERDEVICENOTASSOCIATED",
            "securityquestions": None, "sessionID": None,
            "smsOtp": None, "token": None, "uid": None, "userIP": None
        }
        raw = self._post_json(self._URL_RBA_OTP, body,
                              referer="https://services.gst.gov.in/services/otpforauth")
        return self._parse_login_response(raw)

    # â”€â”€ Step 3 â€” Download GSTR-1 â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def download_gstr1(self, period: str, year: int,
                       gstin: str = "",
                       force_api: bool = False) -> Gstr1Result:
        """
        Downloads GSTR-1 data.

        Strategy:
          1. Try offline ZIP download (complete data, mirrors C# GetGstr1WithOutOTP).
          2. If ZIP unavailable or force_api=True â†’ download summary + sections via API.

        Args:
            period    : Month "1"â€“"12"
            year      : Calendar year (e.g. 2025 for April, 2026 for Jan)
            gstin     : GSTIN string (used for B2B section in API fallback)
            force_api : Skip ZIP attempt, go directly to per-section API

        Returns:
            Gstr1Result
        """
        result         = Gstr1Result()
        result.period  = period
        result.year    = year
        result.rtn_prd = self._build_rtn_prd(period, year)
        rtn            = result.rtn_prd
        ref            = "https://return.gst.gov.in/returns/auth/gstr1/dashboard"

        # [CRITICAL] Set return period context (Mimics selecting month + Search)
        self._log(f"Setting portal context for {rtn}...")
        
        # Calculate FY string
        prd_int = int(period)
        fy_year = year - 1 if prd_int <= 3 else year
        fy = f"{fy_year}-{str(fy_year+1)[2:]}"
        
        # 1. rtnPrdSelection
        ctx_url = f"{self._BASE_RETURN}/returns/auth/api/rtnPrdSelection?rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1"
        self._get_json(ctx_url, referer=ref)
        
        # 2. dropdown
        drop_url = f"{self._BASE_RETURN}/returns/auth/api/dropdown?rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1"
        self._get_json(drop_url, referer=ref)

        # 3. rolestatus (Mandatory to "open" the month)
        role_url = f"{self._BASE_RETURN}/returns/auth/api/rolestatus?rtn_prd={rtn}&fy={fy}"
        self._get_json(role_url, referer=ref)

        # 4. ustatus (Initializes the return period)
        ustatus_url = f"{self._BASE_RETURN}/returns/auth/api/ustatus?rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1"
        self._get_json(ustatus_url, referer=ref)

        self._log(f"GSTR-1 download: period={period}, year={year}, rtn_prd={rtn}")

        # Warm up
        self._post_login_nav(period, year, fy)

        # formdetails warm-up â€” initialises portal session for this return/period
        # (same pattern as GSTR-2A; without this, all data API calls return empty)
        self._formdetails(rtn, ref, fy)

        # Always fetch summary (compact overview)
        self._log("Fetching GSTR-1 Dashboard summary...")
        sum_url = (f"{self._BASE_RETURN}/returns/auth/api/gstr1/summary"
                   f"?rtn_prd={rtn}&smrytyp=L&fy={fy}")
        summary_raw = self._get_json(sum_url, referer=ref)
        result.summary_json = summary_raw
        try:
            self._last_summary_data = json.loads(summary_raw)
        except:
            self._last_summary_data = {}
        
        # Include profile if already fetched
        if self._profile_data:
            result.ustatus_json = json.dumps(self._profile_data)

        if not force_api:
            # Try to get names from summary to inject into ZIP/API data
            name_map = self._get_ctin_name_map(rtn, ref)

            # Strategy 1: Offline ZIP download
            self._log("Attempting offline ZIP download...")
            zip_jsons = self._download_offline_zip(rtn, ref)
            if zip_jsons:
                result.source    = "zip"
                # Inject names into ZIP JSONs
                result.zip_jsons = self._inject_names_into_jsons(zip_jsons, name_map)
                self._log(f"ZIP download complete â€” {len(zip_jsons)} file(s).")
                return result
            self._log("ZIP download not available â€” falling back to API.")

        # Strategy 2: Per-section API
        self._log("Downloading via per-section API...")
        result.source = "api"

        # B2B â€” per supplier
        # Per-supplier sections
        # Supplier-level sections (paginated per-ctin)
        sec_configs = [
            ("B2B",   "b2b",   "SU"), ("B2BA",  "b2ba",  "SU"),
            ("CDNR",  "cdnr",  "SU"), ("CDNRA", "cdnra", "SU")
        ]
        for sec_name, sec_key, uploaded_by in sec_configs:
            if self._yearly_stop: break
            self._log(f"Fetching {sec_name} suppliers...")
            data_json, count = self._fetch_b2b_section(rtn, ref, gstin, sec_name=sec_name, fy=fy, uploaded_by=uploaded_by)
            result.sections[sec_key] = data_json
            if sec_key == "b2b": result.b2b_count = count
            if count > 0:
                self._log(f"{sec_name}: {count} supplier(s).")
            time.sleep(0.5)

        # Flat sections
        for sec_name, sec_key, uploaded_by in GSTR1_FLAT_SECTIONS:
            if self._yearly_stop: break
            if sec_key in result.sections: continue # skip if already done
            self._log(f"Fetching {sec_name}...")
            data = self._fetch_flat_section(rtn, sec_name, sec_key, uploaded_by, ref, fy=fy)
            result.sections[sec_key] = data
            time.sleep(0.3)

        self._log(f"GSTR-1 API download complete â€” sections: {list(result.sections.keys())}")
        return result

    def trigger_offline_gen(self, period: str, year: int):
        """
        Triggers offline ZIP generation without waiting for it to finish.
        Used to 'pre-heat' the portal for yearly downloads.
        """
        rtn = self._build_rtn_prd(period, year)
        ref = "https://return.gst.gov.in/returns/auth/gstr1"
        self._log(f"Pre-triggering ZIP for {rtn}...")
        
        # Calculate FY string
        prd_int = int(period)
        fy_year = year - 1 if prd_int <= 3 else year
        fy = f"{fy_year}-{str(fy_year+1)[2:]}"
        
        # Essential context calls
        ctx_url = f"{self._BASE_RETURN}/returns/auth/api/rtnPrdSelection?rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1"
        self._get_json(ctx_url, referer=ref)
        ustatus_url = f"{self._BASE_RETURN}/returns/auth/api/ustatus?rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1"
        self._get_json(ustatus_url, referer=ref)
        
        # Trigger generation
        gen_url = (f"{self._BASE_RETURN}/returns/auth/api/offline/download/generate"
                   f"?flag=1&rtn_prd={rtn}&rtn_typ=GSTR1")
        if fy: gen_url += f"&fy={fy}"
        self._session.get(gen_url, headers={
            "Accept":  "application/json, text/plain, */*",
            "Referer": "https://return.gst.gov.in/returns/auth/gstr/offlinedownload",
            **self._common_headers()
        }, verify=False, timeout=15)

    # â”€â”€ Step 4 â€” Logout â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def logout(self):
        try:
            self._session.get(self._URL_LOGOUT, headers={
                "Accept":  "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Referer": "https://services.gst.gov.in/services/auth/fowelcome",
                **self._common_headers()
            }, verify=False, timeout=self._timeout)
            self._log("Logged out.")
        except Exception as ex:
            self._log(f"Logout error (ignored): {ex}")
        finally:
            if not self._session_external:
                self._session.cookies.clear()

    # â”€â”€ Internal â€” offline ZIP download â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _download_offline_zip(self, rtn: str, ref: str) -> list:
        """
        Attempts to download the offline ZIP from the portal.
        Returns list of JSON strings extracted from ZIP file(s).
        Returns [] if not available.

        Mirrors C# GetGstr1WithOutOTP() offline download block.
        """
        try:
            # Check if a recent file is available
            check_url = (f"{self._BASE_RETURN}/returns/auth/api/offline/download/generate"
                         f"?flag=0&rtn_prd={rtn}&rtn_typ=GSTR1")
            check_resp = self._get_json(check_url,
                                        referer="https://return.gst.gov.in/returns/auth/gstr/offlinedownload")

            if not check_resp:
                return []

            has_existing = ('"status":1' in check_resp and
                            "You have downloaded the file last on" in check_resp)

            if not has_existing:
                # Trigger fresh generation (async â€” portal may take seconds/minutes)
                self._log(f"Triggering offline file generation for {rtn}...")

                # Calculate FY string
                prd_int = int(rtn[:2])
                fy_year = int(rtn[2:]) - 1 if prd_int <= 3 else int(rtn[2:])
                fy = f"{fy_year}-{str(fy_year+1)[2:]}"

                nav_url = f"https://return.gst.gov.in/returns/auth/gstr/offlinedownload?rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1"
                self._session.get(nav_url, headers={**self._common_headers(), "Referer": "https://return.gst.gov.in/returns/auth/gstr1"}, verify=False, timeout=self._timeout)

                gen_url = (f"{self._BASE_RETURN}/returns/auth/api/offline/download/generate"
                           f"?flag=1&rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1")
                gen_resp = self._session.get(gen_url, headers={
                    "Accept":  "application/json, text/plain, */*",
                    "Referer": nav_url,
                    **self._common_headers()
                }, verify=False, timeout=self._timeout)

                # If portal explicitly says ZIP not available (e.g. not filed / no data),
                # bail out immediately rather than polling
                gen_text = gen_resp.text if gen_resp else ""
                if gen_text and ('"status":0' in gen_text or
                                 "not eligible" in gen_text.lower() or
                                 "no data" in gen_text.lower()):
                    self._log("ZIP not available for this period â€” using API mode.")
                    return []

                # Poll for ZIP readiness â€” max 3 attempts, 5s apart (~15s total)
                found_zip = False
                for attempt in range(1, 4):
                    stop = getattr(self, "_yearly_stop", False)
                    if stop:
                        self._log("Stop requested during ZIP wait.")
                        return []
                    self._log(f"Waiting for ZIP generation (attempt {attempt}/3)...")
                    time.sleep(5)
                    check_resp = self._get_json(check_url, referer=nav_url)
                    if check_resp and '"status":1' in check_resp and rtn in check_resp:
                        self._log("ZIP file is now available.")
                        found_zip = True
                        break

                if not found_zip:
                    self._log("ZIP not ready after 15s â€” falling back to API mode.")
                    return []

            # Parse number of files
            file_count = 1
            try:
                obj = json.loads(check_resp)
                urls = obj.get("data", {}).get("url", [])
                if urls:
                    file_count = len(urls)
            except Exception:
                pass

            # Download each file
            jsons = []
            for i in range(1, file_count + 1):
                if getattr(self, "_yearly_stop", False):
                    self._log("Stop requested â€” aborting ZIP download.")
                    return []
                
                # Parse available URLs from response
                returned_urls = []
                try:
                    obj = json.loads(check_resp)
                    returned_urls = obj.get("data", {}).get("url", [])
                except: pass
                
                if returned_urls and i-1 < len(returned_urls):
                    rel = returned_urls[i-1]
                    dl_url = f"{self._BASE_RETURN}{rel}" if rel.startswith("/") else rel
                else:
                    dl_url = (f"{self._BASE_RETURN}/returns/auth/api/offline/download/url"
                              f"?rtn_prd={rtn}&rtn_typ=GSTR1&file_num={i}")
                self._log(f"Downloading ZIP file {i}/{file_count}...")
                
                try:
                    # Use a larger timeout for the actual ZIP download (60s connect, 90s read)
                    # and don't use stream=True to avoid iter_content hangs on slow portals
                    resp = self._session.get(dl_url, headers={
                        "Accept":  "application/zip,application/octet-stream,*/*",
                        "Referer": "https://return.gst.gov.in/returns/auth/gstr/offlinedownload",
                        **self._common_headers()
                    }, verify=False, timeout=(30, 60))

                    if resp.status_code != 200:
                        self._log(f"ZIP download returned HTTP {resp.status_code} â€” using API.")
                        return []
                    
                    # Verify content type â€” portal sometimes returns HTML error as 200
                    ct = resp.headers.get("Content-Type", "").lower()
                    if "html" in ct:
                        self._log("Portal returned HTML instead of ZIP â€” using API mode.")
                        return []

                    content = resp.content
                    if not content or len(content) < 100:
                        self._log("ZIP download returned too little content â€” using API mode.")
                        return []

                    # Extract JSON from ZIP
                    try:
                        with zipfile.ZipFile(io.BytesIO(content)) as zf:
                            for name in zf.namelist():
                                if name.lower().endswith(".json"):
                                    jsons.append(zf.read(name).decode("utf-8"))
                    except zipfile.BadZipFile:
                        # Sometimes it's a raw JSON (not zipped) despite the URL
                        try:
                            text = content.decode("utf-8", errors="replace").strip()
                            if text.startswith("{") and text.endswith("}"):
                                jsons.append(text)
                            else:
                                self._log("Bad ZIP file signature â€” using API mode.")
                                return []
                        except:
                            self._log("Bad ZIP file signature â€” using API mode.")
                            return []
                except requests.exceptions.Timeout:
                    self._log("ZIP download connection timed out â€” using API mode.")
                    return []
                except Exception as ex:
                    self._log(f"ZIP download error: {ex} â€” using API mode.")
                    return []

            return jsons

        except Exception as ex:
            self._log(f"Offline ZIP error: {ex}")
            return []

    def _get_ctin_name_map(self, rtn: str, ref: str) -> dict:
        """Extracts GSTIN -> Trade Name map from the already fetched dashboard summary."""
        name_map = {}
        if not hasattr(self, "_last_summary_data") or not self._last_summary_data:
            return name_map
            
        try:
            data = self._last_summary_data.get("data", {})
            # Scan all sections in the summary for counterparty names
            for sec in data.get("sec_sum", []):
                for party in sec.get("cpty_sum", []):
                    ctin = party.get("ctin")
                    # Try all common portal name keys
                    name = party.get("trdnm") or party.get("trade_name") or party.get("trdNm") or party.get("trade_nm")
                    if ctin and name:
                        name_map[ctin] = name
        except Exception as ex:
            self._log(f"Error harvesting names from summary: {ex}")
        
        if not name_map:
            self._log("No names found in dashboard summary. Counterparty names may be missing in Excel.")
        else:
            self._log(f"Harvested {len(name_map)} names from summary.")
            
        return name_map

    def _inject_names_into_jsons(self, jsons: list, name_map: dict) -> list:
        if not name_map: return jsons
        out = []
        for js in jsons:
            try:
                obj = json.loads(js)
                modified = False
                # Inject into both B2B and CDNR
                for sec in ["b2b", "cdnr", "b2ba", "cdnra"]:
                    for party in obj.get(sec, []):
                        ctin = party.get("ctin")
                        if ctin in name_map and not party.get("trdnm"):
                            party["trdnm"] = name_map[ctin]
                            modified = True
                if modified:
                    out.append(json.dumps(obj, ensure_ascii=False))
                else:
                    out.append(js)
            except:
                out.append(js)
        return out

    # â”€â”€ Internal â€” per-section API fallback â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _fetch_b2b_section(self, rtn: str, ref: str, gstin: str = "", sec_name: str = "B2B", fy: str = "", uploaded_by: str = "SU"):
        """
        Fetches a per-supplier section via paginated API:
          1. Try secsummary and totalsummarycount for all possible sub-sections
          2. invoice (paginated) per ctin â†’ get all invoices
        """
        ctins = []
        
        # 1. Sections to search in
        trial_sections = [sec_name]
        if sec_name == "B2B": trial_sections.extend(["B2B_4A", "B2B_4B", "B2B_6C", "B2B_SEZWP", "B2B_SEZWOP"])
        elif sec_name == "B2BA": trial_sections.extend(["B2BA_4A", "B2BA_4B", "B2BA_6C", "B2BA_SEZWP", "B2BA_SEZWOP"])
        elif sec_name == "CDNR": trial_sections.extend(["CDNR_4A", "CDNR_4B", "CDNR_6C", "CDNR_SEZWP", "CDNR_SEZWOP"])
        elif sec_name == "CDNRA": trial_sections.extend(["CDNRA_4A", "CDNRA_4B", "CDNRA_6C", "CDNRA_SEZWP", "CDNRA_SEZWOP"])

        for ts in trial_sections:
            if self._yearly_stop: break
            
            # A. Try secsummary (Full list if small, or first page)
            try:
                sec_url = (f"{self._BASE_RETURN}/returns/auth/api/gstr1/secsummary"
                           f"?rtn_prd={rtn}&sec_name={ts}")
                if fy: sec_url += f"&fy={fy}"
                section_ref = f"{self._BASE_RETURN}/returns/auth/gstr1/{ts.lower()}"
                raw = self._get_json(sec_url, referer=section_ref)
                if raw:
                    obj = json.loads(raw)
                    data = obj.get("data", obj)
                    found = []
                    for key in ("cpty", "cpty_sum", "ctin_list", "b2b"):
                        if key in data and isinstance(data[key], list):
                            found = [e.get("ctin") or e.get("stin") or "" for e in data[key] if isinstance(e, dict)]
                            if found: break
                    if found:
                        ctins.extend([c for c in found if c])
                    else:
                        # Sometimes counterparties are in sec_sum of the response
                        for s in data.get("sec_sum", []):
                            if s.get("sec_nm") == ts or s.get("sec_nm") == sec_name:
                                found_s = [e.get("ctin") or e.get("stin") or "" for e in s.get("cpty_sum", [])]
                                ctins.extend([c for c in found_s if c])
                elif raw and '"status":0' in raw:
                    # If pending failed, try a different approach later
                    pass
            except: pass

            # B. Try totalsummarycount (Paginated list)
            page = 1
            while page < 50:
                if self._yearly_stop: break
                cnt_url = (f"{self._BASE_RETURN}/returns/auth/api/gstr1/totalsummarycount"
                           f"?rtn_prd={rtn}&sec_name={ts}&pageNum={page}")
                if fy: cnt_url += f"&fy={fy}"
                section_ref = f"{self._BASE_RETURN}/returns/auth/gstr1/{ts.lower()}"
                raw = self._get_json(cnt_url, referer=section_ref)
                if not raw or '"status":0' in raw: break
                try:
                    obj = json.loads(raw)
                    data = obj.get("data", obj)
                    p_found = []
                    for key in ("cpty", "ctins", "ctin_list", "cpty_sum", "b2b"):
                        if key in data and isinstance(data[key], list):
                            p_found = [e.get("ctin") or e.get("stin") or "" for e in data[key] if isinstance(e, dict)]
                            if p_found: break
                    if not p_found: break
                    ctins.extend([c for c in p_found if c])
                    if len(p_found) < 10: break
                    page += 1
                except: break

        # 2. Unique ctins
        ctins = list(dict.fromkeys(ctins))

        # 3. Last resort fallback from Dashboard Summary
        if not ctins and hasattr(self, "_last_summary_data"):
            try:
                sec_sums = self._last_summary_data.get("data", {}).get("sec_sum", [])
                for s in sec_sums:
                    sn = s.get("sec_nm", "")
                    if sn == sec_name or sn.startswith(f"{sec_name}_"):
                        cpty_list = s.get("cpty_sum") or s.get("cptySum") or s.get("ctinSum")
                        if cpty_list and isinstance(cpty_list, list):
                            ctins.extend([e.get("ctin", "") for e in cpty_list if e.get("ctin")])
            except: pass

        if not ctins:
            return json.dumps({sec_name.lower(): []}), 0

        ctins = list(dict.fromkeys(ctins))
        supplier_jsons = []
        name_map = self._get_ctin_name_map(rtn, ref)

        def _fetch_party_worker(ctin):
            inv_page = 1
            party_combined = None

            # Try specific sub-sections if base fails
            trial_sections = [sec_name]
            if sec_name == "B2B": trial_sections.extend(["B2B_4A", "B2B_4B", "B2B_6C"])
            elif sec_name == "CDNR": trial_sections.extend(["CDNR_4A", "CDNR_4B", "CDNR_6C"])

            for trial_sec in trial_sections:
                inv_page = 1
                seen_inums = set()  # deduplicate across pages
                section_ref = f"{self._BASE_RETURN}/returns/auth/gstr1/{trial_sec.lower()}/invoice/proc"

                while inv_page < 50:
                    inv_url = (f"{self._BASE_RETURN}/returns/auth/api/gstr1/invoice"
                               f"?ctin={ctin}&rtn_prd={rtn}&sec_name={trial_sec}&pageNum={inv_page}")
                    if uploaded_by:
                        inv_url += f"&uploaded_by={uploaded_by}"
                    if fy: inv_url += f"&fy={fy}"
                    
                    raw = self._get_json(inv_url, referer=section_ref)
                    
                    # If SU fails, try without it (common for filed returns)
                    # Support multiple keys: inv, processedInvoice, nt, ntA
                    has_data = raw and any(k in raw for k in ['"inv"', '"processedInvoice"', '"nt"', '"ntA"', '"b2b"'])
                    if not has_data and (uploaded_by == "SU" or "uploaded_by=SU" in inv_url):
                        inv_url_no_su = inv_url.replace("&uploaded_by=SU", "")
                        raw = self._get_json(inv_url_no_su, referer=section_ref)

                    if not raw or not any(k in raw for k in ['"inv"', '"b2b"', '"processedInvoice"', '"cdnr"', '"nt"', '"ntA"']):
                        break

                    try:
                        obj = json.loads(raw)
                        inner = obj.get("data", obj)
                        if not isinstance(inner, dict): break

                        page_records = []
                        rec_key = "inv"
                        for k in ["inv", "b2b", "processedInvoice", "cdnr", "nt", "cdnra"]:
                            if k in inner and isinstance(inner[k], list):
                                page_records = inner[k]
                                rec_key = k
                                break

                        # Deduplicate: filter out any records whose inum/nt_num we've already seen
                        new_records = []
                        for rec in page_records:
                            key = rec.get("inum") or rec.get("nt_num") or rec.get("val")
                            if key not in seen_inums:
                                seen_inums.add(key)
                                new_records.append(rec)

                        # If no new records on this page, the portal is repeating â€” stop
                        if not new_records:
                            break

                        if party_combined is None:
                            party_combined = inner
                            if "ctin" not in party_combined: party_combined["ctin"] = ctin
                            if ctin in name_map and not party_combined.get("trdnm"):
                                party_combined["trdnm"] = name_map[ctin]
                            party_combined[rec_key] = new_records
                        else:
                            if rec_key not in party_combined: party_combined[rec_key] = []
                            party_combined[rec_key].extend(new_records)

                        if len(page_records) < 10: break
                        inv_page += 1
                    except: break

                # Break if we found any data for this supplier
                if party_combined and any(k in party_combined for k in ["inv", "b2b", "processedInvoice", "nt", "ntA"]):
                    break
            return party_combined

        # Use ThreadPoolExecutor for concurrent fetching
        max_workers = 5  # Safe limit for GST portal
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_ctin = {executor.submit(_fetch_party_worker, ctin): ctin for ctin in ctins}
            for future in as_completed(future_to_ctin):
                party_data = future.result()
                if party_data:
                    supplier_jsons.append(party_data)

        return json.dumps({sec_name.lower(): supplier_jsons}, ensure_ascii=False), len(supplier_jsons)

    def _fetch_flat_section(self, rtn: str, sec_name: str, sec_key: str, uploaded_by: Optional[str], ref: str, fy: str = "") -> str:
        """Fetches a flat (non-per-supplier) section with pagination support."""
        all_data = []
        template_obj = None
        list_key = None
        
        seen_sigs = set()
        page = 1
        while page < 50:
            url = (f"{self._BASE_RETURN}/returns/auth/api/gstr1/invoice"
                   f"?rtn_prd={rtn}&sec_name={sec_name}&pageNum={page}")
            if uploaded_by:
                url += f"&uploaded_by={uploaded_by}"
            if fy:
                url += f"&fy={fy}"
            if sec_name == "NIL":
                url = (f"{self._BASE_RETURN}/returns/auth/api/gstr1/invoice"
                       f"?inum=NIL&rtn_prd={rtn}&sec_name=NIL&fy={fy}")
            
            section_ref = f"{ref}/{sec_name.lower()}"
            raw = self._get_json(url, referer=section_ref)
            
            # If SU fails, try without it
            if (not raw or '"status":0' in raw) and uploaded_by == "SU":
                url_no_su = url.replace("&uploaded_by=SU", "")
                raw = self._get_json(url_no_su, referer=section_ref)

            if not raw or not raw.strip().endswith("}"):
                break
            
            try:
                obj = json.loads(raw)
                inner = obj.get("data", obj)
                
                # Identify the list to aggregate (e.g. 'b2cs', 'inv', 'itms')
                found_list = None
                found_key = None
                
                if isinstance(inner, list):
                    found_list = inner
                elif isinstance(inner, dict):
                    # Prioritize key matches
                    for k in [sec_key, sec_key.lower(), sec_name, sec_name.lower(), "inv", "itms", "doc_det", "hsn", "doc_issue", "cpty", "itms_det"]:
                        if k in inner and isinstance(inner[k], list):
                            found_list = inner[k]
                            found_key = k
                            break
                    # Fallback to any list found in data
                    if found_list is None:
                        for k, v in inner.items():
                            if isinstance(v, list) and k not in ["sec_sum"]:
                                found_list = v
                                found_key = k
                                break
                
                if found_list is None:
                    if page == 1: return raw 
                    break
                
                if page == 1:
                    template_obj = obj
                    list_key = found_key
                
                # Deduplicate or stop if portal repeats same records
                new_records = []
                for rec in found_list:
                    # Using a subset of keys for signature to be robust
                    sig = json.dumps(rec, sort_keys=True)
                    if sig not in seen_sigs:
                        seen_sigs.add(sig)
                        new_records.append(rec)
                
                if not new_records and page > 1:
                    # Portal returned already seen records (misbehaving pagination)
                    break
                
                all_data.extend(new_records)
                if len(found_list) < 10: break
                page += 1
                if sec_name == "NIL": break
            except: break
        
        if template_obj and all_data:
            # Reconstruct the template with the full aggregated list
            if list_key:
                if "data" in template_obj:
                    template_obj["data"][list_key] = all_data
                else:
                    template_obj[list_key] = all_data
            return json.dumps(template_obj, ensure_ascii=False)
            
        return "{}"

    # â”€â”€ Internal â€” shared â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _formdetails(self, rtn: str, referer: str, fy: str = ""):
        url = (f"{self._BASE_RETURN}/returns/auth/api/formdetails"
               f"?rtn_prd={rtn}&rtn_typ=GSTR1")
        if fy: url += f"&fy={fy}"
        self._get_json(url, referer=referer)

    def _build_rtn_prd(self, period: str, year: int) -> str:
        return period.zfill(2) + str(year)

    def _post_login_nav(self, period: str = "", year: int = 0, fy: str = ""):
        self._log("Post-login navigation...")
        h = {"Accept": "application/json, text/plain, */*", **self._common_headers()}
        
        # 1. Dashboard (Main)
        self._session.get("https://return.gst.gov.in/returns/auth/dashboard", headers={
            **h, "Referer": "https://services.gst.gov.in/services/auth/dashboard"
        }, verify=False, timeout=self._timeout)

        # 2. Return Dashboard
        self._session.get("https://return.gst.gov.in/returns/auth/gstr/dashboard", headers={
            **h, "Referer": "https://return.gst.gov.in/returns/auth/dashboard"
        }, verify=False, timeout=self._timeout)

        if period and year:
            rtn = self._build_rtn_prd(period, year)
            if not fy:
                p_int = int(period)
                fy_y = year - 1 if p_int <= 3 else year
                fy = f"{fy_y}-{str(fy_y+1)[2:]}"

            # 3. Search return period
            ctx_url = f"https://return.gst.gov.in/returns/auth/api/rtnPrdSelection?rtn_prd={rtn}&fy={fy}&rtn_typ=GSTR1"
            self._get_json(ctx_url, referer="https://return.gst.gov.in/returns/auth/gstr/dashboard")

            # 4. rolestatus
            role_url = f"https://return.gst.gov.in/returns/auth/api/rolestatus?rtn_prd={rtn}&fy={fy}"
            self._get_json(role_url, referer="https://return.gst.gov.in/returns/auth/gstr/dashboard")
        
        self._log("Navigation done.")

    def _get_json(self, url: str, referer: str = "") -> str:
        headers = {"Accept": "application/json, text/plain, */*",
                   "Referer": referer, **self._common_headers()}
        try:
            resp = self._session.get(url, headers=headers,
                                     verify=False, timeout=self._timeout)
            if resp.status_code != 200:
                self._log(f"API Warning [{resp.status_code}] for {url[:60]}...")
            return resp.text
        except Exception as ex:
            self._log(f"GET error [{url[:60]}]: {ex}")
            return ""

    def _post_json(self, url: str, body: dict, referer: str = "") -> str:
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json;charset=utf-8",
            "Referer": referer, "Sec-Fetch-Dest": "empty",
            **self._common_headers()
        }
        resp = self._session.post(url, data=json.dumps(body, separators=(",", ":")),
                                  headers=headers, verify=False, timeout=self._timeout)
        return resp.text

    def _common_headers(self) -> dict:
        return {
            "User-Agent":     USER_AGENT,
            "Sec-Fetch-Mode": "no-cors",
            "Sec-Fetch-Site": "same-origin",
            "Cache-Control":  "no-cache",
        }

    def _parse_login_response(self, raw: str) -> LoginResult:
        try:
            data = json.loads(raw)
            if "successCode" in data:
                return LoginResult(LoginResult.SUCCESS, "Login successful.")
            if "errorCode" in data:
                code = data["errorCode"]
                msg  = GST_ERRORS.get(code, f"GST portal error: {code}")
                if code == "RSK_1000":
                    return LoginResult(LoginResult.OTP_REQUIRED,
                                       "OTP required. Call login_with_otp(otp).")
                return LoginResult(LoginResult.FAILED, msg)
        except Exception as ex:
            self._log(f"Parse login error: {ex}")
        return LoginResult(LoginResult.FAILED, "Unexpected portal response.")

    def _extract_device_id(self, raw: str):
        try:
            obj = json.loads(raw)
            did = obj.get("deviceID", "")
            if did:
                self._device_id = str(did)
        except Exception:
            pass

    def _setup_logging(self):
        self._logger = logging.getLogger(f"Gstr1Downloader.{id(self)}")
        self._logger.setLevel(logging.DEBUG)
        if not self._logger.handlers:
            fmt = logging.Formatter("%(asctime)s  %(message)s", datefmt="%d-%m-%Y %H:%M:%S")
            ch = logging.StreamHandler()
            ch.setLevel(logging.INFO)
            ch.setFormatter(fmt)
            self._logger.addHandler(ch)
            if self._log_path:
                Path(self._log_path).parent.mkdir(parents=True, exist_ok=True)
                fh = logging.FileHandler(self._log_path, encoding="utf-8")
                fh.setLevel(logging.DEBUG)
                fh.setFormatter(fmt)
                self._logger.addHandler(fh)

    def _log(self, msg: str):
        self._logger.debug(msg)
        if self._log_callback:
            try:
                self._log_callback(msg)
            except Exception:
                pass


# =============================================================================
# Save helper
# =============================================================================

def save_gstr1(result: Gstr1Result, output_dir: str, username: str) -> list:
    """
    Saves GSTR-1 data to file(s).
    ZIP mode: GSTR1_{username}_{year}_{period}[_{n}].json
    API mode: GSTR1_{username}_{year}_{period}_sections.json
    Returns list of saved paths.
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    prd   = result.period.zfill(2)
    paths = []

    if result.source == "zip" and result.zip_jsons:
        for i, js in enumerate(result.zip_jsons):
            if len(result.zip_jsons) == 1:
                name = f"GSTR1_{username}_{result.year}_{prd}.json"
            else:
                name = f"GSTR1_{username}_{result.year}_{prd}_{i+1}.json"
            path = Path(output_dir) / name
            path.write_text(js, encoding="utf-8")
            print(f"  Saved: {path}")
            paths.append(str(path))
    else:
        # API mode â€” save combined file
        combined = {
            "rtn_prd":     result.rtn_prd,
            "period":      result.period,
            "year":        result.year,
            "source":      result.source,
            "summary":     result.summary_json,
            "sections":    result.sections,
            "profile":     json.loads(result.ustatus_json) if result.ustatus_json else {},
            "b2b_count":   result.b2b_count,
        }
        name = f"GSTR1_{username}_{result.year}_{prd}_sections.json"
        path = Path(output_dir) / name
        path.write_text(json.dumps(combined, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"  Saved: {path}")
        paths.append(str(path))

    return paths


# =============================================================================
# CLI
# =============================================================================

def main():
    import os
    print("=" * 55)
    print("  GSTR-1 Downloader â€” GST Portal Direct")
    print("=" * 55)

    OUTPUT_DIR   = r"D:\CompuOffice Online\compugst\GSTR1_Output"
    CAPTCHA_FILE = os.path.join(OUTPUT_DIR, "captcha.png")
    LOG_FILE     = os.path.join(OUTPUT_DIR, "gstr1_log.txt")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    username = input("GST Username    : ").strip()
    password = input("GST Password    : ").strip()
    period   = input("Period (1-12)   : ").strip()
    year     = int(input("Year (e.g. 2024): ").strip())
    gstin    = input("GSTIN (optional): ").strip()

    d = Gstr1Downloader(log_path=LOG_FILE)

    print("\n[1] Fetching captcha...")
    b64 = d.get_captcha_base64()
    Path(CAPTCHA_FILE).write_bytes(base64.b64decode(b64))
    print(f"    Captcha saved: {CAPTCHA_FILE}")
    captcha = input("Captcha text    : ").strip()

    print("\n[2] Logging in...")
    r = d.login(username, password, captcha)
    if r.otp_required:
        otp = input("    OTP: ").strip()
        r = d.login_with_otp(otp)
    if not r.is_success:
        print(f"Login failed: {r.message}")
        return

    print(f"\n[3] Downloading GSTR-1 â€” {period}/{year}...")
    result = d.download_gstr1(period, year, gstin=gstin)
    print(f"    Source: {result.source}")
    save_gstr1(result, OUTPUT_DIR, username)

    print("\n[4] Logging out...")
    d.logout()
    print("Done.")


if __name__ == "__main__":
    main()
