"""
GSTR-2B Downloader — Python Automation
========================================
Direct HTTP automation against the GST portal.
No GSP. No browser/Selenium. Pure requests + cookies.

Install dependencies:
    pip install requests Pillow

Files created by this script:
    D:/CompuOffice Online/compugst/gstr2b_downloader.py   ← this file
    D:/CompuOffice Online/compugst/Gstr2BDownloader.cs    ← C# version
    D:/CompuOffice Online/compugst/Gstr2BDownloader_Usage.cs

Author  : Generated from CompuOffice GST source analysis
Portal  : https://gstr2b.gst.gov.in
"""

import os
import io
import json
import random
import base64
import logging
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ─────────────────────────────────────────────────────────────────────────────
# Optional: show captcha image in terminal or save to file
# ─────────────────────────────────────────────────────────────────────────────
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


# =============================================================================
# Constants
# =============================================================================

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)

# Browser fingerprint — sent in login POST body, required by GST portal
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
            "Flash": "20,0,0,306",
            "WindowsMediaPlayer": "12,0,7601,17514",
            "VBVersion": "10.0.16521",
            "ConnectionType": "lan",
            "AddressBook": "6,1,7601,17514",
            "BrowsingPack": "10,0,9200,16521",
            "DHTMLDataBinding": "10,0,9200,16521",
            "IEHelp": "10,0,9200,16521",
            "IEHelpEngine": "6,2,9200,16521",
            "OfflineBrowsingPack": "10,0,9200,16521",
            "OutlookExpress": "6,1,7601,17514",
            "WindowsDektopUpdate": "6,1,7601,17514"
        },
        "NetscapePlugins": {},
        "Screen": {
            "FullHeight": 768, "AvlHeight": 724,
            "FullWidth": 1366, "AvlWidth": 1366,
            "BufferDepth": 0, "ColorDepth": 24, "PixelDepth": 24,
            "DeviceXDPI": 96, "DeviceYDPI": 96,
            "FontSmoothing": True, "UpdateInterval": 0
        },
        "System": {
            "Platform": "Win32", "OSCPU": "x86",
            "userLanguage": "en-IN", "Timezone": -330
        }
    },
    "ExternalIP": "",
    "MESC": {"mesc": "mi=2;cd=150;id=30;mesc=207943;mesc=224290"}
}, separators=(',', ':'))

# GST portal error code → human-readable message
GST_ERRORS = {
    "SWEB_9000": "Invalid Captcha. Please try again.",
    "AUTH_9002": "Invalid UserId or Password. Please try again.",
    "AUTH_9033": "Password has Expired. Please reset your password.",
    "SWEB_8000": "Error at GSTN site. Please try after sometime.",
    "SWEB_9014": "Account locked after 3 wrong password attempts. Reset password first.",
    "RSK_1000" : "OTP Required",
    "SWEB_9003": "Wrong OTP entered.",
}


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


class Gstr2BResult:
    def __init__(self):
        self.json_chunks: list[str] = []   # one entry per portal file chunk
        self.file_count: int        = 0    # fc value from portal
        self.mode: str              = "M"  # M = Monthly, Q = Quarterly

    @property
    def combined_json(self) -> str:
        """Chunks joined with ~ (same format as CompuOffice DB storage)."""
        return "~".join(self.json_chunks)

    def __repr__(self):
        return (f"Gstr2BResult(chunks={len(self.json_chunks)}, "
                f"fc={self.file_count}, mode={self.mode!r})")


# =============================================================================
# Main downloader class
# =============================================================================

class Gstr2BDownloader:
    """
    GST portal GSTR-2B automation.

    Usage:
        d = Gstr2BDownloader(log_path="gstr2b.log")

        # 1. Get captcha
        captcha_b64 = d.get_captcha_base64()
        # → save to PNG, show to user, collect their input

        # 2. Login
        result = d.login("username", "password", "captcha_text")
        if result.otp_required:
            otp = input("Enter OTP: ")
            result = d.login_with_otp(otp)

        # 3. Download
        data = d.download_gstr2b(period="3", year=2024, mode="M")
        for i, chunk in enumerate(data.json_chunks):
            with open(f"gstr2b_part{i+1}.json", "w") as f:
                f.write(chunk)

        # 4. Logout
        d.logout()
    """

    # ── GST portal base URLs ──────────────────────────────────────────────────
    _URL_CAPTCHA      = "https://services.gst.gov.in/services/captcha"
    _URL_AUTHENTICATE = "https://services.gst.gov.in/services/authenticate"
    _URL_RBA_OTP      = "https://services.gst.gov.in/services/validate/rba/otp"
    _URL_USTATUS      = "https://services.gst.gov.in/services/api/ustatus"
    _URL_DROPDOWN     = "https://return.gst.gov.in/returns/auth/api/dropdown"
    _URL_LOGOUT       = "https://services.gst.gov.in/services/logout"
    _URL_DASHBOARD    = "https://services.gst.gov.in/services/auth/api/dashboard/itcashldg"

    _URL_2B_GETDATA   = "https://gstr2b.gst.gov.in/gstr2b/auth/api/gstr2b/getdata"
    _URL_2B_GETJSON   = "https://gstr2b.gst.gov.in/gstr2b/auth/api/gstr2b/getjson"
    _URL_2BQ_GETJSON  = "https://gstr2b.gst.gov.in/gstr2b/auth/api/gstr2bq/getjson"

    _REFERER_MONTHLY   = "https://gstr2b.gst.gov.in/gstr2b/auth/gstr2bdwld"
    _REFERER_QUARTERLY = "https://gstr2b.gst.gov.in/gstr2b/auth/gstr2bqdwld"

    def __init__(self, log_path: str = "", timeout: int = 120,
                 session=None):
        """
        Args:
            log_path: File path for debug logs. Empty string = no logging.
            timeout : Request timeout in seconds (default 120).
            session : Optional requests.Session to share (e.g. from combined GUI).
                      When provided the caller owns the session lifecycle.
        """
        self._log_path = log_path
        self._timeout  = timeout
        self._device_id: str    = ""
        self._last_login_raw: str = ""

        if session is not None:
            self._session = session
            self._session_external = True
        else:
            self._session = requests.Session()
            self._session_external = False
            retry = Retry(
                total=5,
                backoff_factor=1,
                status_forcelist=[500, 502, 503, 504],
                allowed_methods=["GET", "POST"]
            )
            adapter = HTTPAdapter(max_retries=retry)
            self._session.mount("https://", adapter)
            self._session.mount("http://",  adapter)

        # Disable SSL verification warnings (portal uses its own chain)
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        self._setup_logging()
        self._log("Gstr2BDownloader initialized.")

    # ─────────────────────────────────────────────────────────────────────────
    # STEP 1 — Get captcha
    # ─────────────────────────────────────────────────────────────────────────

    def get_captcha_base64(self) -> str:
        """
        Fetches the captcha image from the GST portal login page.

        Returns:
            Base64-encoded PNG string.
            Convert with base64.b64decode(result) to get raw bytes → save as .png

        Note:
            Resets the cookie session on every call (fresh login required).
        """
        if not self._session_external:
            self._session.cookies.clear()
        self._device_id = ""
        self._last_login_raw = ""

        url = f"{self._URL_CAPTCHA}?rnd={random.random()}"
        self._log(f"Fetching captcha: {url}")

        headers = {
            "Accept":            "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
            "Referer":           "https://services.gst.gov.in/services/login",
            "X-Requested-With":  "XMLHttpRequest",
            "Sec-Fetch-Dest":    "image",
            "Sec-Fetch-Mode":    "no-cors",
            "Sec-Fetch-Site":    "same-origin",
            "Cache-Control":     "no-cache",
            "User-Agent":        USER_AGENT,
        }

        resp = self._session.get(url, headers=headers, verify=False, timeout=self._timeout)
        resp.raise_for_status()

        img_b64 = base64.b64encode(resp.content).decode("utf-8")
        self._log(f"Captcha fetched ({len(resp.content)} bytes).")
        return img_b64

    def save_captcha_image(self, output_path: str) -> str:
        """
        Convenience method: fetches captcha and saves it as a PNG file.

        Args:
            output_path: Where to save the PNG (e.g. "captcha.png")

        Returns:
            Same output_path (for chaining).
        """
        b64 = self.get_captcha_base64()
        img_bytes = base64.b64decode(b64)
        Path(output_path).write_bytes(img_bytes)
        self._log(f"Captcha saved to: {output_path}")
        return output_path

    def show_captcha(self, output_path: str = "captcha.png") -> str:
        """
        Fetches captcha, saves as PNG, and opens it for the user.

        Returns:
            The output file path.
        """
        self.save_captcha_image(output_path)
        if PIL_AVAILABLE:
            try:
                img = Image.open(output_path)
                img.show()
            except Exception:
                pass
        else:
            # Fallback: open with OS default viewer
            os.startfile(output_path) if os.name == "nt" else None
        return output_path

    # ─────────────────────────────────────────────────────────────────────────
    # STEP 2 — Login
    # ─────────────────────────────────────────────────────────────────────────

    def login(self, username: str, password: str, captcha_text: str) -> LoginResult:
        """
        Logs in to the GST portal.

        Args:
            username    : GST portal username
            password    : Portal password (max 15 chars)
            captcha_text: Text read from the captcha image

        Returns:
            LoginResult with status SUCCESS, OTP_REQUIRED, or FAILED.
            If OTP_REQUIRED → call login_with_otp(otp) next.
        """
        if len(password) > 15:
            return LoginResult(
                LoginResult.FAILED,
                "Password too long — GST portal allows maximum 15 characters."
            )

        body = {
            "username": username,
            "password": password,
            "captcha":  captcha_text,
            "mFP":      MFP_JSON,
            "deviceID": self._device_id if self._device_id else None,
            "type":     "username"
        }

        self._log(f"Logging in as: {username}")
        raw = self._post_json(
            url     = self._URL_AUTHENTICATE,
            body    = body,
            referer = "https://services.gst.gov.in/services/login"
        )
        self._last_login_raw = raw
        self._log(f"Login response: {raw[:300]}")

        # Save deviceID for RBA OTP call
        self._extract_device_id(raw)

        return self._parse_login_response(raw)

    # ─────────────────────────────────────────────────────────────────────────
    # STEP 2b — OTP (only when login returns OTP_REQUIRED)
    # ─────────────────────────────────────────────────────────────────────────

    def login_with_otp(self, otp: str) -> LoginResult:
        """
        Submits OTP for Risk-Based Authentication.
        Call this only after login() returns LoginResult.OTP_REQUIRED.

        Args:
            otp: OTP sent to the taxpayer's registered mobile/email.
        """
        device_id = self._device_id
        try:
            prev = json.loads(self._last_login_raw)
            if prev.get("deviceID"):
                device_id = str(prev["deviceID"])
        except Exception:
            pass

        body = {
            "username":          None,
            "password":          None,
            "captcha":           None,
            "mFP":               MFP_JSON,
            "deviceID":          device_id,
            "type":              "username",
            "applIP":            None,
            "email":             None,
            "emailOtp":          None,
            "mobileNo":          None,
            "oidar":             None,
            "otpAuthSts":        "true",
            "otpDetails":        json.dumps({"otp": otp}),
            "refNum":            None,
            "riskLvl":           "INCREASEAUTH",
            "role":              None,
            "ruleAnnotation":    (
                "UNKNOWNUSER=N;MFPMISMATCH=N;NEGATIVECOUNTRY=N;"
                "UNTRUSTEDIP=N;UNKNOWNDEVICEID=N;DEVICEVELOCITY=N;"
                "ZONEHOPPING=N;USERDEVICENOTASSOCIATED=Y;GST_USER_VELOCITY=N;"
            ),
            "ruleMatched":       "USERDEVICENOTASSOCIATED",
            "securityquestions": None,
            "sessionID":         None,
            "smsOtp":            None,
            "token":             None,
            "uid":               None,
            "userIP":            None
        }

        self._log("Submitting OTP for RBA...")
        raw = self._post_json(
            url     = self._URL_RBA_OTP,
            body    = body,
            referer = "https://services.gst.gov.in/services/otpforauth"
        )
        self._log(f"OTP response: {raw[:300]}")
        return self._parse_login_response(raw)

    # ─────────────────────────────────────────────────────────────────────────
    # STEP 3 — Download GSTR-2B
    # ─────────────────────────────────────────────────────────────────────────

    def download_gstr2b(
        self,
        period: str,
        year: int,
        mode: str = "M",
        max_retries: int = 5
    ) -> Gstr2BResult:
        """
        Downloads GSTR-2B JSON data from the GST portal.

        Must be called after a successful login().

        Args:
            period     : Month number as string "1"–"12" (e.g. "3" for March)
            year       : Calendar year (e.g. 2025 for April, 2026 for Jan)
            mode       : "M" = Monthly (default), "Q" = Quarterly (QRMP scheme)
            max_retries: Max retries per request on partial/garbled response

        Returns:
            Gstr2BResult with json_chunks list (one entry per file chunk).

        Example:
            result = d.download_gstr2b("3", 2024, "M")
            # result.json_chunks[0] → full GSTR-2B JSON from portal
            # result.file_count     → fc value (0 = single file)
        """
        result          = Gstr2BResult()
        result.mode     = mode
        endpoint_idx    = 0   # 0 = primary URL, 1 = fallback quarterly URL

        # Calculate FY string
        prd_int = int(period)
        fy_year = year - 1 if prd_int <= 3 else year
        fy = f"{fy_year}-{str(fy_year+1)[2:]}"

        self._log(f"GSTR-2B download: period={period}, year={year}, mode={mode}, fy={fy}")

        # Warm up session on return portal
        self._post_login_navigation(period, year, fy)

        # ── Phase 1: metadata call (fc = 0) ───────────────────────────────────
        self._log(f"Phase 1: metadata call for period={period}, year={year}, mode={mode}")
        initial_json = self._fetch_with_retry(
            period=period, year=year, fc=0,
            mode=mode, endpoint_idx=endpoint_idx,
            max_retries=max_retries, fy=fy
        )
        if not initial_json:
            raise RuntimeError(
                "GST portal not providing proper data. Please try again after some time."
            )

        # Parse file count
        fc = self._parse_fc(initial_json)
        result.file_count = fc
        self._log(f"Portal file count (fc) = {fc}")

        if fc == 0:
            # Single-file response — initial JSON is the complete data
            result.json_chunks.append(initial_json)
            return result

        # ── Phase 2: download each file chunk ─────────────────────────────────
        self._log(f"Phase 2: downloading {fc} chunk(s)...")
        for i in range(1, fc + 1):
            self._log(f"Fetching chunk {i}/{fc}")
            chunk = self._fetch_with_retry(
                period=period, year=year, fc=i,
                mode=mode, endpoint_idx=endpoint_idx,
                max_retries=max_retries
            )
            if not chunk:
                raise RuntimeError(f"Failed to download chunk {i}/{fc} after {max_retries} retries.")
            result.json_chunks.append(chunk)

        return result

    # ─────────────────────────────────────────────────────────────────────────
    # STEP 4 — Logout
    # ─────────────────────────────────────────────────────────────────────────

    def logout(self):
        """Logs out from the GST portal and clears the cookie session."""
        try:
            headers = {
                "Accept":  "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Referer": "https://services.gst.gov.in/services/auth/fowelcome",
                **self._common_headers()
            }
            self._session.get(
                self._URL_LOGOUT,
                headers=headers, verify=False, timeout=self._timeout
            )
            self._log("Logged out.")
        except Exception as ex:
            self._log(f"Logout error (ignored): {ex}")
        finally:
            if not self._session_external:
                self._session.cookies.clear()

    # ─────────────────────────────────────────────────────────────────────────
    # Internal — URL building
    # ─────────────────────────────────────────────────────────────────────────

    def _build_url(self, period: str, year: int, fc: int, mode: str, endpoint_idx: int, fy: str = "") -> str:
        """
        Builds the correct GSTR-2B portal URL.

        Period/Year encoding (matches CompuOffice original):
          Months 1,2,3 → url_year = year + 1   (Jan/Feb/Mar belong to next calendar year)
          Months 4–12  → url_year = year
        """
        rtnprd   = period.zfill(2) + str(year)
        fy_param = f"&fy={fy}" if fy else ""

        if fc > 0:
            # ── Chunk download ────────────────────────────────────────────────
            if mode == "Q":
                if endpoint_idx == 1:
                    return f"{self._URL_2B_GETJSON}?rtnprd={rtnprd}{fy_param}&fn={fc}&quart=Y"
                else:
                    return f"{self._URL_2BQ_GETJSON}?itcprd={rtnprd}{fy_param}&fn={fc}"
            return f"{self._URL_2B_GETJSON}?rtnprd={rtnprd}{fy_param}&fn={fc}&quart=N"
        else:
            # ── Metadata call ─────────────────────────────────────────────────
            if mode == "Q":
                if endpoint_idx == 1:
                    return f"{self._URL_2B_GETJSON}?rtnprd={rtnprd}{fy_param}&quart=Y"
                else:
                    return f"{self._URL_2BQ_GETJSON}?itcprd={rtnprd}{fy_param}"
            return f"{self._URL_2B_GETDATA}?rtnprd={rtnprd}{fy_param}&quart=N"

    def _get_referer(self, mode: str) -> str:
        return self._REFERER_QUARTERLY if mode == "Q" else self._REFERER_MONTHLY

    # ─────────────────────────────────────────────────────────────────────────
    # Internal — fetch with retry logic
    # ─────────────────────────────────────────────────────────────────────────

    def _fetch_with_retry(
        self, period: str, year: int, fc: int,
        mode: str, endpoint_idx: int, max_retries: int, fy: str = ""
    ) -> Optional[str]:
        """
        Fetches a single 2B endpoint with up to max_retries retries.
        Handles:
          - Incomplete responses (not ending with "}")
          - Error responses ("error" in body)
          - Quarterly summary mismatch (m1sum/m1cnt)
        """
        retries = 0
        while retries <= max_retries:
            try:
                url     = self._build_url(period, year, fc, mode, endpoint_idx, fy)
                referer = self._get_referer(mode)

                self._log(f"GET {url}")
                headers = {
                    "Accept":   "application/json, text/plain, */*",
                    "Referer":  referer,
                    **self._common_headers()
                }
                resp = self._session.get(
                    url, headers=headers,
                    verify=False, timeout=self._timeout
                )
                text = resp.text

                self._log(
                    f"Response [{len(text)} chars]: {text[:200]}"
                )

                if not text.rstrip().endswith("}"):
                    self._log(f"Incomplete response, retry {retries + 1}")
                    retries += 1
                    time.sleep(1)
                    continue

                if '"error"' in text:
                    self._log(f"Error in response, switching endpoint. Retry {retries + 1}")
                    retries += 1
                    endpoint_idx = 1   # try fallback quarterly endpoint
                    time.sleep(1)
                    continue

                if "m1sum" in text.lower() or "m1cnt" in text.lower():
                    self._log(f"Quarterly summary (m1sum/m1cnt), switching endpoint. Retry {retries + 1}")
                    retries += 1
                    endpoint_idx = 1
                    time.sleep(1)
                    continue

                return text   # ← good response

            except Exception as ex:
                self._log(f"Request error, retry {retries + 1}: {ex}")
                retries += 1
                time.sleep(2)

        return None  # exhausted retries

    # ─────────────────────────────────────────────────────────────────────────
    # Internal — post-login navigation
    # ─────────────────────────────────────────────────────────────────────────

    def _post_login_navigation(self, period: str = "", year: int = 0, fy: str = ""):
        """
        Warms up the session on services.gst.gov.in and return.gst.gov.in.
        Required before accessing gstr2b.gst.gov.in.
        """
        self._log("Post-login navigation...")

        common = self._common_headers()

        # 1. User status
        self._session.get(
            self._URL_USTATUS,
            headers={
                "Accept":  "application/json, text/plain, */*",
                "Referer": "https://services.gst.gov.in/services/auth/dashboard",
                **common
            },
            verify=False, timeout=self._timeout
        )

        # 2. Return portal dropdown
        drop_url = self._URL_DROPDOWN
        if fy: drop_url += f"?fy={fy}"
        self._session.get(
            drop_url,
            headers={
                "Accept":  "application/json, text/plain, */*",
                "Referer": "https://return.gst.gov.in/returns/auth/dashboard",
                **common
            },
            verify=False, timeout=self._timeout
        )

        if period and year:
            rtn = period.zfill(2) + str(year)
            # Use base domain from dropdown URL
            base = "https://return.gst.gov.in"
            role_url = f"{base}/returns/auth/api/rolestatus?rtn_prd={rtn}"
            if fy: role_url += f"&fy={fy}"
            self._session.get(
                role_url,
                headers={
                    "Accept":  "application/json, text/plain, */*",
                    "Referer": "https://return.gst.gov.in/returns/auth/dashboard",
                    **common
                },
                verify=False, timeout=self._timeout
            )
        self._log("Post-login navigation done.")

    # ─────────────────────────────────────────────────────────────────────────
    # Internal — HTTP helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _post_json(self, url: str, body: dict, referer: str = "") -> str:
        headers = {
            "Accept":       "application/json, text/plain, */*",
            "Content-Type": "application/json;charset=utf-8",
            "Referer":      referer,
            "Sec-Fetch-Dest": "empty",
            **self._common_headers()
        }
        resp = self._session.post(
            url,
            data=json.dumps(body, separators=(",", ":")),
            headers=headers,
            verify=False,
            timeout=self._timeout
        )
        return resp.text

    def _common_headers(self) -> dict:
        return {
            "User-Agent":     USER_AGENT,
            "Sec-Fetch-Mode": "no-cors",
            "Sec-Fetch-Site": "same-origin",
            "Cache-Control":  "no-cache",
        }

    # ─────────────────────────────────────────────────────────────────────────
    # Internal — parse helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _parse_login_response(self, raw: str) -> LoginResult:
        try:
            data = json.loads(raw)
            if "successCode" in data:
                return LoginResult(LoginResult.SUCCESS, "Login successful.")

            if "errorCode" in data:
                code = data["errorCode"]
                msg  = GST_ERRORS.get(code, f"GST portal error: {code}")
                if code == "RSK_1000" or msg == "OTP Required":
                    return LoginResult(
                        LoginResult.OTP_REQUIRED,
                        "OTP required (Risk-Based Auth). "
                        "An OTP has been sent to the registered mobile/email. "
                        "Call login_with_otp(otp) to continue."
                    )
                return LoginResult(LoginResult.FAILED, msg)

        except Exception as ex:
            self._log(f"_parse_login_response error: {ex}")

        return LoginResult(LoginResult.FAILED, "Unexpected response from portal. Check logs.")

    def _parse_fc(self, json_str: str) -> int:
        """Extract data.fc (file count) from portal JSON response."""
        try:
            obj = json.loads(json_str)
            fc  = obj.get("data", {}).get("fc", 0)
            return int(fc)
        except Exception:
            return 0

    def _extract_device_id(self, raw: str):
        try:
            obj = json.loads(raw)
            did = obj.get("deviceID", "")
            if did:
                self._device_id = str(did)
        except Exception:
            pass

    # ─────────────────────────────────────────────────────────────────────────
    # Internal — logging
    # ─────────────────────────────────────────────────────────────────────────

    def _setup_logging(self):
        self._logger = logging.getLogger("Gstr2BDownloader")
        self._logger.setLevel(logging.DEBUG)
        fmt = logging.Formatter("%(asctime)s  %(message)s", datefmt="%d-%m-%Y %H:%M:%S")

        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(fmt)
        self._logger.addHandler(ch)

        # File handler (if path given)
        if self._log_path:
            Path(self._log_path).parent.mkdir(parents=True, exist_ok=True)
            fh = logging.FileHandler(self._log_path, encoding="utf-8")
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(fmt)
            self._logger.addHandler(fh)

    def _log(self, msg: str):
        self._logger.debug(msg)


# =============================================================================
# Helpers — saving results
# =============================================================================

def save_result(result: Gstr2BResult, output_dir: str, codeno: str, period: str, year: int):
    """
    Saves all chunks from a Gstr2BResult to JSON files.

    File naming matches CompuOffice convention:
        GSTR2B_Return_{codeno}_{year}_{period}.json          (single file)
        GSTR2B_Return_{codeno}_{year}_{period}_{i}.json      (chunk i)
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    prd = period.zfill(2)

    for i, chunk in enumerate(result.json_chunks):
        if len(result.json_chunks) == 1:
            filename = f"GSTR2B_Return_{codeno}_{year}_{prd}.json"
        else:
            filename = f"GSTR2B_Return_{codeno}_{year}_{prd}_{i + 1}.json"

        path = os.path.join(output_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(chunk)
        print(f"  Saved: {path}")

    print(f"  Total chunks: {len(result.json_chunks)} | fc={result.file_count}")


# =============================================================================
# CLI — run directly
# =============================================================================

def main():
    """
    Interactive command-line usage.
    Run: python gstr2b_downloader.py
    """
    print("=" * 60)
    print("  GSTR-2B Downloader — GST Portal Automation")
    print("=" * 60)

    # ── Config ────────────────────────────────────────────────────────────────
    OUTPUT_DIR = r"D:\CompuOffice Online\compugst\GSTR2B_Output"
    LOG_FILE   = os.path.join(OUTPUT_DIR, "gstr2b_log.txt")
    CAPTCHA_FILE = os.path.join(OUTPUT_DIR, "captcha.png")
    # ─────────────────────────────────────────────────────────────────────────

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    username = input("GST Username    : ").strip()
    password = input("GST Password    : ").strip()
    period   = input("Period (1-12)   : ").strip()
    year     = int(input("Year (e.g. 2024): ").strip())
    mode     = input("Mode M/Q [M]    : ").strip().upper() or "M"
    codeno   = input("Code No         : ").strip()

    downloader = Gstr2BDownloader(log_path=LOG_FILE, timeout=120)

    # ── Step 1: Captcha ───────────────────────────────────────────────────────
    print("\n[1] Fetching captcha...")
    downloader.show_captcha(CAPTCHA_FILE)
    print(f"    Captcha image saved: {CAPTCHA_FILE}")
    print("    Open the image, read the text, then enter it below.")
    captcha_text = input("Captcha text    : ").strip()

    # ── Step 2: Login ─────────────────────────────────────────────────────────
    print("\n[2] Logging in...")
    login_result = downloader.login(username, password, captcha_text)
    print(f"    Status: {login_result.status} — {login_result.message}")

    if login_result.otp_required:
        otp = input("    OTP received on mobile/email: ").strip()
        login_result = downloader.login_with_otp(otp)
        print(f"    OTP status: {login_result.status} — {login_result.message}")

    if not login_result.is_success:
        print("Login failed. Exiting.")
        return

    # ── Step 3: Download GSTR-2B ──────────────────────────────────────────────
    print(f"\n[3] Downloading GSTR-2B — period={period}, year={year}, mode={mode}...")
    try:
        result = downloader.download_gstr2b(period, year, mode)
        print(f"    Downloaded {len(result.json_chunks)} chunk(s), fc={result.file_count}")
        save_result(result, OUTPUT_DIR, codeno, period, year)
    except Exception as ex:
        print(f"    Download error: {ex}")

    # ── Step 4: Logout ────────────────────────────────────────────────────────
    print("\n[4] Logging out...")
    downloader.logout()
    print("Done.")


# =============================================================================
# Bulk download example — multiple parties, one period
# =============================================================================

def bulk_download_example():
    """
    Example: download GSTR-2B for multiple parties for the same period.
    Each party needs its own captcha entry (GST portal requires it per login).
    """
    OUTPUT_DIR = r"D:\CompuOffice Online\compugst\GSTR2B_Output"
    LOG_FILE   = os.path.join(OUTPUT_DIR, "bulk_log.txt")
    PERIOD     = "3"
    YEAR       = 2024
    MODE       = "M"

    parties = [
        {"codeno": "C001", "username": "user1", "password": "pass1"},
        {"codeno": "C002", "username": "user2", "password": "pass2"},
        # add more...
    ]

    for party in parties:
        print(f"\n{'=' * 50}")
        print(f"Processing: {party['codeno']}")
        print(f"{'=' * 50}")

        downloader = Gstr2BDownloader(log_path=LOG_FILE, timeout=120)

        try:
            # Captcha per party
            captcha_file = os.path.join(OUTPUT_DIR, f"captcha_{party['codeno']}.png")
            downloader.save_captcha_image(captcha_file)
            print(f"Captcha saved: {captcha_file}")
            captcha_text = input(f"Enter captcha for {party['codeno']}: ").strip()

            # Login
            result = downloader.login(party["username"], party["password"], captcha_text)

            if result.otp_required:
                otp = input(f"OTP for {party['codeno']}: ").strip()
                result = downloader.login_with_otp(otp)

            if not result.is_success:
                print(f"Login failed: {result.message}")
                continue

            # Download
            data = downloader.download_gstr2b(PERIOD, YEAR, MODE)
            save_result(data, OUTPUT_DIR, party["codeno"], PERIOD, YEAR)

        except Exception as ex:
            print(f"Error for {party['codeno']}: {ex}")
        finally:
            downloader.logout()


if __name__ == "__main__":
    main()
