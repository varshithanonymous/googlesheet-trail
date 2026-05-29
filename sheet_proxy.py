#!/usr/bin/env python3
"""
Local static server + minimal proxy for fetching published Google Sheets CSV.

Why:
- Browsers block cross-origin reads from docs.google.com (CORS).
- Public "CORS proxy" services are unreliable and often blocked.

Usage:
  cd E:\downloads\project
  python sheet_proxy.py
  # then open:
  # http://127.0.0.1:8080/CRM_source_code.html
"""

from __future__ import annotations

import mimetypes
import os
import urllib.parse
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
import base64
import subprocess
import urllib.error
import re


ALLOWED_HOSTS = {
    "docs.google.com",
    "spreadsheets.google.com",
}


def search_ddg(query: str) -> list[dict[str, str]]:
    """
    Search DuckDuckGo Lite and return top results (title, url, snippet).
    """
    from bs4 import BeautifulSoup
    import urllib.request
    import urllib.parse

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    url = "https://lite.duckduckgo.com/lite/"
    data = urllib.parse.urlencode({"q": query}).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")

    results = []
    try:
        proxies = urllib.request.getproxies() or {}
        opener = urllib.request.build_opener(urllib.request.ProxyHandler(proxies))
        print(f"[Search] Executing DuckDuckGo Lite search for query: {query}", flush=True)
        with opener.open(req, timeout=20) as response:
            html = response.read()
            soup = BeautifulSoup(html, "html.parser")

            tables = soup.find_all("table")
            if len(tables) >= 3:
                # The third table typically contains search results
                table = tables[2]
                rows = table.find_all("tr")
                current_result = {}
                for row in rows:
                    td = row.find("td", class_="result-snippet")
                    if td:
                        if current_result:
                            current_result["snippet"] = td.get_text().strip()
                            results.append(current_result)
                            current_result = {}
                    else:
                        a = row.find("a", class_="result-link")
                        if a:
                            current_result = {
                                "title": a.get_text().strip(),
                                "url": a.get("href", ""),
                            }

            # Fallback if structure is different
            if not results:
                links = soup.find_all("a", class_="result-link")
                for link in links:
                    results.append({
                        "title": link.get_text().strip(),
                        "url": link.get("href", ""),
                        "snippet": ""
                    })
        print(f"[Search] Success! Found {len(results)} search results.", flush=True)
    except Exception as e:
        print(f"[Search] DuckDuckGo search failed: {e}", flush=True)
    return results[:12]


class Handler(SimpleHTTPRequestHandler):
    server_version = "SheetProxy/1.0"

    def end_headers(self) -> None:
        # Allow the HTML app to call /proxy from the browser.
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        # Some environments may append trailing slashes or otherwise vary the path.
        if parsed.path.rstrip('/') == "/config":
            self._handle_config()
            return
        if parsed.path.rstrip("/") == "/ai_contacts":
            self._handle_ai_contacts(parsed)
            return
        if parsed.path.rstrip("/") == "/version":
            out = json.dumps({"ok": True, "server": self.server_version}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(out)
            return
        super().do_GET()

    def _handle_proxy(self, parsed: urllib.parse.ParseResult, debug: bool = False) -> None:
        qs = urllib.parse.parse_qs(parsed.query)
        url = (qs.get("url") or [""])[0].strip()
        if not url:
            self.send_error(400, "Missing url")
            return

        try:
            target = urllib.parse.urlparse(url)
        except Exception:
            self.send_error(400, "Invalid url")
            return

        if target.scheme not in ("http", "https"):
            self.send_error(400, "Invalid scheme")
            return

        host = (target.hostname or "").lower()
        if host not in ALLOWED_HOSTS:
            self.send_error(403, "Host not allowed")
            return

        # Fetch upstream (respect system proxy settings if present).
        proxies = urllib.request.getproxies() or {}
        opener = urllib.request.build_opener(urllib.request.ProxyHandler(proxies))

        req = urllib.request.Request(
            url,
            headers={
                # Some Google endpoints behave better with a UA.
                "User-Agent": "Mozilla/5.0 (SheetProxy)",
                "Accept": "text/csv,text/plain,*/*",
            },
            method="GET",
        )

        try:
            with opener.open(req, timeout=20) as resp:
                body = resp.read()
                ctype = resp.headers.get("Content-Type") or "text/plain; charset=utf-8"
                final_url = getattr(resp, "url", url)
                status = getattr(resp, "status", 200)
                headers = dict(resp.headers.items())
        except Exception as e:
            # Some networks allow browser/WinINet but block Python sockets to docs.google.com.
            # Fallback to PowerShell Invoke-WebRequest (uses Windows networking stack).
            try:
                safe_url = url.replace("'", "''")
                ps = [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    (
                        "$ProgressPreference='SilentlyContinue';"
                        f"$u='{safe_url}';"
                        "$r=Invoke-WebRequest -UseBasicParsing -Uri $u -TimeoutSec 25;"
                        "[Console]::OutputEncoding=[Text.Encoding]::UTF8;"
                        "$r.Content"
                    ),
                ]
                out = subprocess.check_output(ps, stderr=subprocess.STDOUT, timeout=30)
                body = out
                ctype = "text/csv; charset=utf-8"
                final_url = url
                status = 200
                headers = {"X-Fallback": "powershell"}
            except Exception as e2:
                self.send_error(502, f"Upstream fetch failed: {e} | PS fallback failed: {e2}")
                return

        if debug:
            payload = {
                "requested_url": url,
                "final_url": final_url,
                "status": status,
                "content_type": ctype,
                "length": len(body),
                "headers": headers,
                "sample_b64": base64.b64encode(body[:400]).decode("ascii"),
            }
            out = json.dumps(payload, indent=2).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(out)
            return

        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _handle_config(self) -> None:
        """Return JSON with credentials from environment variables."""
        # Define the env variable names expected for configuration
        api_key = os.getenv("OPENAI_API_KEY", "").strip()
        sheet_url = os.getenv("GOOGLE_SHEET_URL", "").strip()
        config = {"openai_api_key": api_key, "sheet_url": sheet_url}
        out = json.dumps(config).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(out)
        return

    def _handle_ai_contacts(self, parsed: urllib.parse.ParseResult) -> None:
        qs = urllib.parse.parse_qs(parsed.query)
        q = (qs.get("q") or [""])[0].strip()
        if not q:
            self.send_error(400, "Missing q")
            return

        # Get API key from browser first, then fall back to env
        qs_apikey = (qs.get("apikey") or [""])[0].strip()
        env_openai = os.environ.get("OPENAI_API_KEY", "").strip()
        
        # Strictly use OpenAI GPT-4o
        api_key = None
        if qs_apikey:
            api_key = qs_apikey
            print(f"[AI] Using OpenAI from browser settings (key: {api_key[:10]}...)", flush=True)
        elif env_openai:
            api_key = env_openai
            print(f"[AI] Using OpenAI from server environment variable", flush=True)
        else:
            self.send_error(400, "Missing OpenAI API Key - please save it in the CRM Sheet Setup or set the OPENAI_API_KEY environment variable.")
            return

        # Perform the DuckDuckGo Search to fetch top 10+ real-time searches
        search_results = search_ddg(q)
        search_results_str = ""
        for idx, res in enumerate(search_results, 1):
            search_results_str += f"{idx}. {res['title']}\n   URL: {res['url']}\n   Snippet: {res.get('snippet', '')}\n\n"

        # Build payload with GPT-4o, 2500 max_tokens, and json_object response format
        payload = {
            "model": "gpt-4o",
            "temperature": 0.4,
            "max_tokens": 2500,
            "response_format": {"type": "json_object"},
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "You are an expert lead generation assistant. You are given a user search query and real search results. "
                        "Identify exactly 10 real companies/leads matching the query. "
                        "Extract and enrich the details for these 10 real companies/leads using the search results and your knowledge. "
                        "For each contact, provide a contact name (realistic or real executive/founder/manager name if available), "
                        "their designation (job title like CEO, Founder, Manager), company name, company size (e.g. 10-50, 50-200, 1000+), "
                        "country, city, phone (real public phone or realistic business phone), email (real public email or realistic business email, e.g. info@company.com), "
                        "linkedin/website URL (use the real company website from search results or their LinkedIn), services (what they do), "
                        "address (exact or general location), source (set to 'other'), and status (set to 'neutral').\n\n"
                        "You MUST return ONLY valid JSON matching this exact schema:\n"
                        "{ \"contacts\": [ {"
                        "\"name\":\"\", \"des\":\"\", \"company\":\"\", \"size\":\"\", \"country\":\"\", \"city\":\"\", "
                        "\"phone\":\"\", \"email\":\"\", \"linkedin\":\"\", \"services\":\"\", \"address\":\"\", "
                        "\"source\":\"other\", \"status\":\"neutral\", \"followup\":\"\", \"futype\":\"\", "
                        "\"lastcontact\":\"\", \"funotes\":\"\" } ] }"
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"Search Query: {q}\n\n"
                        f"Real Web Search Results:\n{search_results_str}\n"
                        "Extract and generate exactly 10 structured contacts based on these real search results. "
                        "Return ONLY the JSON object, nothing else."
                    ),
                },
            ],
        }

        endpoint = "https://api.openai.com/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "SheetProxy/AI",
        }

        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(endpoint, data=data, headers=headers, method="POST")

        try:
            proxies = urllib.request.getproxies() or {}
            opener = urllib.request.build_opener(urllib.request.ProxyHandler(proxies))
            with opener.open(req, timeout=60) as resp:
                body = resp.read()
        except urllib.error.HTTPError as e:
            msg = e.read().decode("utf-8", errors="replace") if hasattr(e, "read") else str(e)
            self.send_error(502, f"AI HTTPError: {e.code} {msg}")
            return
        except Exception as e:
            self.send_error(502, f"AI request failed: {e}")
            return

        try:
            j = json.loads(body.decode("utf-8", errors="replace"))
            content = (
                j.get("choices", [{}])[0]
                .get("message", {})
                .get("content", "")
            )
        except Exception as e:
            self.send_error(502, f"Bad AI envelope: {e}")
            return

        # Parse the JSON output securely
        try:
            if isinstance(content, (dict, list)):
                out = content
            else:
                s = str(content or "").strip()
                
                # Remove markdown code fences and formatting
                s = re.sub(r"```(?:json)?\s*", "", s)
                s = s.replace("```", "")
                s = s.strip()
                
                try:
                    out = json.loads(s)
                except Exception:
                    # Find first { and match with last }
                    start_idx = s.find("{")
                    if start_idx >= 0:
                        depth = 0
                        for i in range(start_idx, len(s)):
                            if s[i] == "{":
                                depth += 1
                            elif s[i] == "}":
                                depth -= 1
                                if depth == 0:
                                    json_str = s[start_idx:i+1]
                                    try:
                                        out = json.loads(json_str)
                                    except Exception:
                                        json_str = json_str.replace("\\n", " ").replace("\n", " ")
                                        out = json.loads(json_str)
                                    break
                        else:
                            raise ValueError("Unmatched braces in JSON")
                    else:
                        raise ValueError("No JSON object found in model output")
        except Exception as e:
            sample = str(content or "")[:400].replace("\n", "\\n")
            self.send_error(502, f"Bad AI response JSON: {e}. Sample: {sample}")
            return

        resp_bytes = json.dumps(out).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(resp_bytes)


def main() -> None:
    port = int(os.environ.get("PORT", "8080"))
    # Serve from the directory where this script lives.
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    httpd = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Serving on http://127.0.0.1:{port}")
    httpd.serve_forever()


if __name__ == "__main__":
    main()
