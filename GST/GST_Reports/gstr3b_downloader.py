"""
GSTR-3B Downloader — Python Automation
========================================
Direct HTTP against the GST portal (no GSP, no Selenium).
Mirrors GstOnlineActivity.GetGSTR3BData() from the C# codebase.

Flow:
  1. get_captcha_base64()
  2. login(username, password, captcha)
     └─ login_with_otp(otp)  if OTP_REQUIRED
  3. download_gstr3b(period, year)
  4. logout()

Install: pip install requests Pillow
"""

import json
import random
import base64
import logging
import time
from pathlib import Path
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ─────────────────────────────────────────────────────────────────────────────
# Constants  (identical to gstr2b_downloader — same portal, same fingerprint)
# ─────────────────────────────────────────────────────────────────────────────

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


class Gstr3BResult:
    def __init__(self):
        self.summary_json: str = ""     # /gstr3b/summary response
        self.payment_json: str = ""     # /gstr3b/taxpayble response
        self.ustatus_json: str = ""     # /ustatus response (contains lgl_nm)
        self.period: str       = ""
        self.year:   int       = 0
        self.rtn_prd: str      = ""     # MMYYYY used in URL

    def __repr__(self):
        return f"Gstr3BResult(period={self.period}, year={self.year}, len={len(self.summary_json)})"


# =============================================================================
# Downloader
# =============================================================================

class Gstr3BDownloader:
    """
    GSTR-3B downloader — direct portal HTTP.
    Mirrors GstOnlineActivity.GetGSTR3BData() from C# codebase.

    Steps:
      1. get_captcha_base64()
      2. login(user, pass, captcha)   → LoginResult
         login_with_otp(otp)          → LoginResult  (only if OTP_REQUIRED)
      3. download_gstr3b(period, year) → Gstr3BResult
      4. logout()
    """

    _URL_CAPTCHA      = "https://services.gst.gov.in/services/captcha"
    _URL_AUTHENTICATE = "https://services.gst.gov.in/services/authenticate"
    _URL_RBA_OTP      = "https://services.gst.gov.in/services/validate/rba/otp"
    _URL_USTATUS      = "https://services.gst.gov.in/services/api/ustatus"
    _URL_DROPDOWN     = "https://return.gst.gov.in/returns/auth/api/dropdown"
    _URL_LOGOUT       = "https://services.gst.gov.in/services/logout"

    _BASE_RETURN      = "https://return.gst.gov.in"

    def __init__(self, log_path: str = "", timeout: int = 120,
                 session: Optional[requests.Session] = None):
        self._timeout       = timeout
        self._log_path      = log_path
        self._device_id     = ""
        self._last_login_raw = ""
        self._profile_data   = {}  # Store fetched profile

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
        self._log("Gstr3BDownloader ready.")

    # ── Step 2.5 — Profile ───────────────────────────────────────────────────

    def fetch_profile(self) -> dict:
        """Fetches company name using the confirmed 'ustatus' API response."""
        self._log("Fetching official company name from portal (ustatus API)...")
        try:
            # Hit the ustatus API which we know contains "bname"
            raw = self._get_json(self._URL_USTATUS, 
                                 referer="https://services.gst.gov.in/services/auth/fowelcome")
            
            if not raw or "{" not in raw:
                self._log("Failed to get a valid response from ustatus API.")
                return {}

            data = json.loads(raw)
            
            # The key is confirmed to be "bname" from the portal network logs
            name = data.get("bname") or data.get("lgl_nm") or data.get("trdnm")
            
            if name:
                self._profile_data = data
                self._log(f"SUCCESS: Company Name captured -> {name}")
                return data
            else:
                self._log(f"FAILED: 'bname' key not found in API response. Available keys: {list(data.keys())}")
                
        except Exception as e:
            self._log(f"Profile API fetch exception: {e}")
        return {}

    # ── Step 1 — Captcha ──────────────────────────────────────────────────────

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

    # ── Step 2 — Login ────────────────────────────────────────────────────────

    def login(self, username: str, password: str, captcha_text: str) -> LoginResult:
        if len(password) > 15:
            return LoginResult(LoginResult.FAILED,
                               "Password too long — GST portal max 15 characters.")
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

    # ── Step 3 — Download GSTR-3B ─────────────────────────────────────────────

    def download_gstr3b(self, period: str, year: int,
                        max_retries: int = 5) -> Gstr3BResult:
        """
        Downloads GSTR-3B summary for the given period/year.

        Mirrors GstOnlineActivity.GetGSTR3BData():
          1. POST-login navigation (warm-up return portal)
          2. GET negativeLiabDetails  (warm-up)
          3. GET formdetails           (warm-up)
          4. GET gstr3b/summary        ← MAIN DATA
          5. GET getTxnStatus          (warm-up)
          6. GET checkLtFeeLiab        (warm-up)

        Args:
            period : Month "1"–"12"
            year   : Calendar year (e.g. 2025 for April, 2026 for Jan)
        """
        result = Gstr3BResult()
        result.period  = period
        result.year    = year
        # Calculate FY string
        prd_int = int(period)
        fy_year = year - 1 if prd_int <= 3 else year
        fy = f"{fy_year}-{str(fy_year+1)[2:]}"

        self._log(f"GSTR-3B download: period={period}, year={year}, rtn_prd={result.rtn_prd}, fy={fy}")

        # Warm up return portal session and capture profile info
        self._post_login_nav(period, year, fy)
        
        # Use already-fetched profile if available, else fetch now
        if self._profile_data:
            result.ustatus_json = json.dumps(self._profile_data)
        else:
            try:
                result.ustatus_json = self._get_json(self._URL_USTATUS, 
                                                     referer="https://services.gst.gov.in/services/auth/dashboard")
            except:
                pass

        rtn = result.rtn_prd
        ref = f"{self._BASE_RETURN}/returns/auth/gstr3b"

        def _get(url):
            if fy:
                sep = "&" if "?" in url else "?"
                url += f"{sep}fy={fy}"
            return self._get_json(url, referer=ref)

        # Step 1: negativeLiabDetails (warm-up only)
        self._log("Step 1: negativeLiabDetails")
        _get(f"{self._BASE_RETURN}/returns/auth/api/gstr3b/negativeLiabDetails?rtn_prd={rtn}&indicator=next")

        # Step 2: formdetails (warm-up only)
        self._log("Step 2: formdetails")
        _get(f"{self._BASE_RETURN}/returns/auth/api/formdetails?rtn_prd={rtn}&rtn_typ=GSTR3B")

        # Step 3: summary — MAIN DATA with retry
        self._log("Step 3: gstr3b/summary (main data)")
        for attempt in range(max_retries + 1):
            text = _get(f"{self._BASE_RETURN}/returns/auth/api/gstr3b/summary?rtn_prd={rtn}")
            if text and text.rstrip().endswith("}") and '"error"' not in text:
                result.summary_json = text
                break
            self._log(f"  Retry {attempt+1}: incomplete response")
            time.sleep(1)
        else:
            raise RuntimeError("GSTR-3B: portal not returning valid summary data.")

        # Step 4: getTxnStatus (warm-up — ignore response)
        self._log("Step 4: getTxnStatus")
        _get(f"{self._BASE_RETURN}/returns/auth/api/gstr3b/getTxnStatus?rtn_prd={rtn}")

        # Step 5: checkLtFeeLiab (warm-up — ignore response)
        self._log("Step 5: checkLtFeeLiab")
        _get(f"{self._BASE_RETURN}/returns/auth/api/gstr3b/checkLtFeeLiab?ret_period={rtn}")

        # Step 6: Payment of Tax Data (Mirrors GetGSTR3BDataPaymentOfTax)
        self._log("Step 6: Fetching Payment of Tax data")
        
        # Navigation to payment page (warm-up)
        pay_ref = f"{self._BASE_RETURN}/returns/auth/gstr3b/payment"
        self._session.get(
            f"{self._BASE_RETURN}/pages/returns/gstr3b/payment/payment.html",
            headers={
                "Accept": "application/json, text/plain, */*",
                "Referer": pay_ref,
                **self._common_headers()
            }, verify=False, timeout=self._timeout
        )

        # taxpayble endpoint
        for attempt in range(max_retries + 1):
            pay_text = self._get_json(
                f"{self._BASE_RETURN}/returns/auth/api/gstr3b/taxpayble?rtn_prd={rtn}",
                referer=pay_ref
            )
            if pay_text and pay_text.rstrip().endswith("}") and '"error"' not in pay_text:
                result.payment_json = pay_text
                break
            self._log(f"  Retry {attempt+1} (payment): incomplete response")
            time.sleep(1)
        
        self._log(f"GSTR-3B download complete — Summary: {len(result.summary_json)}, Payment: {len(result.payment_json)} chars.")
        return result

    # ── Step 4 — Logout ───────────────────────────────────────────────────────

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

    # ── Internal ──────────────────────────────────────────────────────────────

    def _build_rtn_prd(self, period: str, year: int) -> str:
        """MMYYYY."""
        return period.zfill(2) + str(year)

    def _post_login_nav(self, period: str = "", year: int = 0, fy: str = ""):
        """Warm up services + return portal. Mirrors GstAftloginReqUrl()."""
        self._log("Post-login navigation...")
        h = {"Accept": "application/json, text/plain, */*", **self._common_headers()}

        self._session.get(self._URL_USTATUS, headers={
            **h, "Referer": "https://services.gst.gov.in/services/auth/dashboard"
        }, verify=False, timeout=self._timeout)

        drop_url = self._URL_DROPDOWN
        if fy: drop_url += f"?fy={fy}"
        self._session.get(drop_url, headers={
            **h, "Referer": "https://return.gst.gov.in/returns/auth/dashboard"
        }, verify=False, timeout=self._timeout)
        
        if period and year:
            rtn = self._build_rtn_prd(period, year)
            role_url = f"{self._BASE_RETURN}/returns/auth/api/rolestatus?rtn_prd={rtn}"
            if fy: role_url += f"&fy={fy}"
            self._session.get(
                role_url,
                headers={**h, "Referer": "https://return.gst.gov.in/returns/auth/dashboard"},
                verify=False, timeout=self._timeout
            )
        self._log("Navigation done.")

    def _get_json(self, url: str, referer: str = "") -> str:
        headers = {"Accept": "application/json, text/plain, */*",
                   "Referer": referer, **self._common_headers()}
        resp = self._session.get(url, headers=headers,
                                 verify=False, timeout=self._timeout)
        return resp.text

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
        self._logger = logging.getLogger(f"Gstr3BDownloader.{id(self)}")
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


# =============================================================================
# Save helper
# =============================================================================

def save_gstr3b(result: Gstr3BResult, output_dir: str, username: str):
    """Saves GSTR-3B unified JSON (summary + payment) to file."""
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    prd  = result.period.zfill(2)
    name = f"GSTR3B_{username}_{result.year}_{prd}.json"
    path = Path(output_dir) / name

    # Merge into a single JSON structure
    unified_data = {}
    try:
        if result.summary_json:
            unified_data["summary"] = json.loads(result.summary_json)
        if result.payment_json:
            unified_data["payment"] = json.loads(result.payment_json)
        if result.ustatus_json:
            try:
                unified_data["profile"] = json.loads(result.ustatus_json)
            except:
                unified_data["profile_raw"] = result.ustatus_json
    except Exception as e:
        # Fallback to raw text if JSON parsing fails
        unified_data["summary_raw"] = result.summary_json
        unified_data["payment_raw"] = result.payment_json
        unified_data["error"] = str(e)

    path.write_text(json.dumps(unified_data, indent=2), encoding="utf-8")
    print(f"  Saved: {path}")
    return str(path)


# =============================================================================
# CLI
# =============================================================================

def main():
    import os
    print("=" * 55)
    print("  GSTR-3B Downloader — GST Portal Direct")
    print("=" * 55)

    OUTPUT_DIR   = r"D:\CompuOffice Online\compugst\GSTR3B_Output"
    CAPTCHA_FILE = os.path.join(OUTPUT_DIR, "captcha.png")
    LOG_FILE     = os.path.join(OUTPUT_DIR, "gstr3b_log.txt")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    username = input("GST Username    : ").strip()
    password = input("GST Password    : ").strip()
    period   = input("Period (1-12)   : ").strip()
    year     = int(input("Year (e.g. 2024): ").strip())

    d = Gstr3BDownloader(log_path=LOG_FILE)

    print("\n[1] Fetching captcha...")
    b64 = d.get_captcha_base64()
    img_bytes = base64.b64decode(b64)
    Path(CAPTCHA_FILE).write_bytes(img_bytes)
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

    print(f"\n[3] Downloading GSTR-3B — {period}/{year}...")
    result = d.download_gstr3b(period, year)
    save_gstr3b(result, OUTPUT_DIR, username)

    print("\n[4] Logging out...")
    d.logout()
    print("Done.")


if __name__ == "__main__":
    main()
