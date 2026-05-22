"""
GSTR-2A Downloader — Python Automation
========================================
Direct HTTP against the GST portal (no GSP, no Selenium).
Mirrors GstOnlineActivity GSTR-2A section from C# codebase.

Sections downloaded:
  B2B  — per-supplier invoices (counter-party)
  B2BA — per-supplier amended invoices
  CDN  — credit/debit notes per supplier
  CDNA — amended CDN per supplier
  TDS  — TDS deducted (single call)
  TCS  — TCS collected (single call)

Flow:
  1. get_captcha_base64()
  2. login(username, password, captcha)
     └─ login_with_otp(otp)  if OTP_REQUIRED
  3. download_gstr2a(period, year)
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
# Constants
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

_NULL_RESPONSE = '{"status_cd":"0","stackTrace":[],"suppressedExceptions":[]}'
_THRESHOLD_MSG  = "exceeded threshold"
_NO_INVOICE_MSG = "No Invoices found"


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


class Gstr2AResult:
    """
    Holds all downloaded GSTR-2A section data.

    sections dict keys: "b2b", "b2ba", "cdn", "cdna", "tds", "tcs"
    Each value is the raw JSON string from the portal.
    """
    def __init__(self):
        self.sections: dict        = {}   # section_name → JSON string
        self.supplier_counts: dict = {}   # section_name → number of suppliers fetched
        self.warnings: list        = []   # non-fatal messages (threshold errors, etc.)
        self.period: str           = ""
        self.year:   int           = 0
        self.rtn_prd: str          = ""

    def __repr__(self):
        counts = {k: v for k, v in self.supplier_counts.items()}
        return f"Gstr2AResult(period={self.period}, year={self.year}, sections={list(self.sections.keys())}, counts={counts})"


# =============================================================================
# Downloader
# =============================================================================

class Gstr2ADownloader:
    """
    GSTR-2A downloader — direct portal HTTP, no GSP.
    Mirrors GstOnlineActivity GSTR-2A section from C# codebase.
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
        self._log("Gstr2ADownloader ready.")

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

    # ── Step 3 — Download GSTR-2A ─────────────────────────────────────────────

    def download_gstr2a(self, period: str, year: int,
                        sections: Optional[list] = None,
                        progress_callback: Optional[callable] = None) -> Gstr2AResult:
        """
        Downloads GSTR-2A for all specified sections.

        Args:
            period  : Month "1"–"12"
            year    : Calendar year (e.g. 2025 for April, 2026 for Jan)
            sections: List of sections to download. Default = all.
                      Choices: "b2b", "b2ba", "cdn", "cdna", "tds", "tcs"

        Returns:
            Gstr2AResult with sections dict populated.
        """
        if sections is None:
            sections = ["b2b", "b2ba", "cdn", "cdna", "tds", "tcs",
                        "isd", "isda", "tdsa", "impg", "impgsez"]

        result          = Gstr2AResult()
        result.period   = period
        result.year     = year
        result.rtn_prd  = self._build_rtn_prd(period, year)
        rtn             = result.rtn_prd

        # Calculate FY string
        prd_int = int(period)
        fy_year = year - 1 if prd_int <= 3 else year
        fy = f"{fy_year}-{str(fy_year+1)[2:]}"

        self._log(f"GSTR-2A download: period={period}, year={year}, rtn_prd={rtn}, fy={fy}")

        # Initial navigation context
        self._post_login_nav(period, year, fy)
        self._formdetails(rtn, fy)

        ref_base = "https://return.gst.gov.in/returns/auth/gstr2/preview"

        # Per-supplier sections (loop through ctin list)
        per_supplier = {
            "b2b":  ("B2B",  "b2b",  "https://return.gst.gov.in/returns/auth/gstr2/preview/b2bcountersupplier"),
            "b2ba": ("B2BA", "b2ba", "https://return.gst.gov.in/returns/auth/gstr2/preview/b2bacounterpreview"),
            "cdn":  ("CDN",  "cdn",  "https://return.gst.gov.in/returns/auth/gstr2/preview/cdncountersupplier"),
            "cdna": ("CDNA", "cdna", "https://return.gst.gov.in/returns/auth/gstr2/preview/cdnacounterpreview"),
        }

        for key, (sec_name, api_seg, inv_ref) in per_supplier.items():
            if key not in sections:
                continue
            self._log(f"Section {sec_name} — fetching supplier list...")
            self._formdetails(rtn, fy)
            data, count, warns = self._fetch_per_supplier_section(
                rtn, sec_name, api_seg, inv_ref, progress_callback)
            result.sections[key]         = data
            result.supplier_counts[key]  = count
            result.warnings.extend(warns)
            self._log(f"Section {sec_name} done — {count} supplier(s).")

        # Flat sections (single API call)
        flat_sections = {
            "tds":     (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/tds?rtn_prd={rtn}",
                        "https://return.gst.gov.in/returns/auth/gstr2/preview/tds"),
            "tcs":     (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/tcs?rtn_prd={rtn}",
                        "https://return.gst.gov.in/returns/auth/gstr2/preview/tcs"),
            "isd":     (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/isd?rtn_prd={rtn}",
                        "https://return.gst.gov.in/returns/auth/gstr2/preview/isd"),
            "isda":    (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/isda?rtn_prd={rtn}",
                        "https://return.gst.gov.in/returns/auth/gstr2/preview/isda"),
            "tdsa":    (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/tdsa?rtn_prd={rtn}",
                        "https://return.gst.gov.in/returns/auth/gstr2/preview/tdsa"),
            "impg":    (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/impg?rtn_prd={rtn}",
                        "https://return.gst.gov.in/returns/auth/gstr2/preview/impg"),
            "impgsez": (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/impgsez?rtn_prd={rtn}",
                        "https://return.gst.gov.in/returns/auth/gstr2/preview/impgsez"),
        }
        for key, (url, ref) in flat_sections.items():
            if key not in sections:
                continue
            self._log(f"Section {key.upper()} — fetching...")
            self._formdetails(rtn, fy)
            text = self._get_json(url, referer=ref)
            if text and text != _NULL_RESPONSE and _NO_INVOICE_MSG not in text:
                result.sections[key] = text
            else:
                result.sections[key] = "{}"
            self._log(f"Section {key.upper()} done.")

        self._log(f"GSTR-2A download complete — sections: {list(result.sections.keys())}")
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

    # ── Internal — per-supplier section logic ─────────────────────────────────

    def _fetch_per_supplier_section(self, rtn: str, sec_name: str,
                                    api_seg: str, inv_ref: str,
                                    progress_callback: Optional[callable] = None):
        """
        Fetches a per-supplier section (B2B, B2BA, CDN, CDNA):
          1. GET ctin list  →  parse cpty[].stin
          2. For each stin: GET individual supplier data
          3. Combine into {"<api_seg>": [...]}

        Returns (json_str, supplier_count, warnings_list)
        """
        warnings = []
        # Step 1: get ctin list
        ctin_url  = (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/ctin"
                     f"?rtn_prd={rtn}&section_name={sec_name}")
        ctin_resp = self._get_json(ctin_url, referer=inv_ref)

        if not ctin_resp or ctin_resp == _NULL_RESPONSE:
            resp_snippet = (ctin_resp[:100] + "...") if ctin_resp else "Empty"
            warnings.append(f"{sec_name}: No supplier list returned (Resp: {resp_snippet}).")
            return "{}", 0, warnings

        if "RET13509" in ctin_resp:
            return "{}", 0, warnings

        try:
            ctin_obj = json.loads(ctin_resp)
            cpty_list = ctin_obj.get("cpty", [])
        except Exception as ex:
            warnings.append(f"{sec_name}: Could not parse supplier list: {ex}")
            return "{}", 0, warnings

        if not cpty_list:
            return f'{{"{api_seg}":[]}}', 0, warnings

        # Step 2: per-supplier detail
        supplier_data = []
        name_map = {}
        for e in cpty_list:
            stin = e.get("stin")
            if stin:
                # Support various portal keys for party name
                name_map[stin] = (e.get("trdnm") or e.get("lgl_nm") or 
                                  e.get("trade_name") or e.get("tradenm") or 
                                  e.get("cname") or e.get("stin_nm"))
        
        total = len(cpty_list)

        for i, entry in enumerate(cpty_list, 1):
            stin = entry.get("stin", "")
            if not stin:
                continue

            if progress_callback:
                progress_callback(f"  {sec_name}: {i}/{total} - {stin}")

            detail_url = (f"{self._BASE_RETURN}/returns/auth/api/gstr2a/{api_seg}"
                          f"?rtn_prd={rtn}&ctin={stin}")
            raw = self._get_json(detail_url, referer=inv_ref)

            if not raw or raw == _NULL_RESPONSE:
                warnings.append(f"{sec_name} [{stin}]: no response.")
                continue
            if _THRESHOLD_MSG in raw:
                warnings.append(f"{sec_name} [{stin}]: >500 records — use offline download.")
                continue
            if "RET13509" in raw or _NO_INVOICE_MSG in raw:
                continue

            try:
                obj = json.loads(raw)
                # Inject name if missing or null in the detailed response
                if isinstance(obj, dict) and stin in name_map:
                    data_obj = obj.get("data", obj)
                    if isinstance(data_obj, dict):
                        for k, v in data_obj.items():
                            if isinstance(v, list):
                                for p in v:
                                    if isinstance(p, dict):
                                        if not p.get("trdnm") and not p.get("lgl_nm"):
                                            p["trdnm"] = name_map[stin]
                                        if not p.get("ctin"):
                                            p["ctin"] = stin
                                break
                
                # Extract the inner list for combining
                inner = obj
                for _ in range(3):
                    if isinstance(inner, dict):
                        vals = list(inner.values())
                        if vals and isinstance(vals[0], (list, dict)):
                            inner = vals[0]
                            break
                        inner = vals[0] if vals else inner
                    elif isinstance(inner, list):
                        break
                
                if isinstance(inner, list):
                    supplier_data.extend(inner)
                else:
                    supplier_data.append(inner)
            except Exception as ex:
                self._log(f"{sec_name} [{stin}]: JSON parse error: {ex}")
                if raw:
                    supplier_data.append({"raw": raw, "error": str(ex)})

        combined = json.dumps({api_seg: supplier_data}, ensure_ascii=False)
        return combined, len(cpty_list), warnings

    # ── Internal — shared helpers ─────────────────────────────────────────────

    def _build_rtn_prd(self, period: str, year: int) -> str:
        return period.zfill(2) + str(year)

    def _formdetails(self, rtn: str, fy: str = ""):
        self._log("Fetching formdetails context...")
        referer = f"{self._BASE_RETURN}/returns/auth/gstr2a"
        url = (f"{self._BASE_RETURN}/returns/auth/api/formdetails"
               f"?rtn_prd={rtn}&rtn_typ=GSTR2A")
        if fy: url += f"&fy={fy}"
        self._get_json(url, referer=referer)

    def _post_login_nav(self, period: str = "", year: int = 0, fy: str = ""):
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
        try:
            resp = self._session.get(url, headers=headers,
                                     verify=False, timeout=self._timeout)
            return resp.text
        except Exception as ex:
            self._log(f"GET error [{url}]: {ex}")
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
        self._logger = logging.getLogger(f"Gstr2ADownloader.{id(self)}")
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

def save_gstr2a(result: Gstr2AResult, output_dir: str, username: str) -> str:
    """
    Saves GSTR-2A data as a combined JSON file.
    Filename: GSTR2A_{username}_{year}_{period}.json
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    prd  = result.period.zfill(2)
    name = f"GSTR2A_{username}_{result.year}_{prd}.json"
    path = Path(output_dir) / name

    combined = {
        "rtn_prd":       result.rtn_prd,
        "period":        result.period,
        "year":          result.year,
        "sections":      result.sections,
        "warnings":      result.warnings,
        "supplier_counts": result.supplier_counts,
    }
    path.write_text(json.dumps(combined, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  Saved: {path}")
    return str(path)


# =============================================================================
# CLI
# =============================================================================

def main():
    import os
    print("=" * 55)
    print("  GSTR-2A Downloader — GST Portal Direct")
    print("=" * 55)

    OUTPUT_DIR   = r"D:\CompuOffice Online\compugst\GSTR2A_Output"
    CAPTCHA_FILE = os.path.join(OUTPUT_DIR, "captcha.png")
    LOG_FILE     = os.path.join(OUTPUT_DIR, "gstr2a_log.txt")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    username = input("GST Username    : ").strip()
    password = input("GST Password    : ").strip()
    period   = input("Period (1-12)   : ").strip()
    year     = int(input("Year (e.g. 2024): ").strip())

    d = Gstr2ADownloader(log_path=LOG_FILE)

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

    print(f"\n[3] Downloading GSTR-2A — {period}/{year}...")
    result = d.download_gstr2a(period, year)

    if result.warnings:
        print(f"\n  Warnings ({len(result.warnings)}):")
        for w in result.warnings:
            print(f"    ! {w}")

    save_gstr2a(result, OUTPUT_DIR, username)

    print("\n[4] Logging out...")
    d.logout()
    print("Done.")


if __name__ == "__main__":
    main()
