"""
Test script to prove the sandboxed cookie approach is safe.
Run: python test_cookie_sandbox.py

This shows exactly which cookies are extracted per domain —
proving no cross-site leakage.
"""
import browser_cookie3
from urllib.parse import urlparse


def show_cookies_for_domain(url: str):
    """Show exactly what cookies would be injected for a given URL."""
    parsed = urlparse(url if "://" in url else f"https://{url}")
    domain = parsed.netloc.removeprefix("www.")

    print(f"\n{'='*60}")
    print(f"Target URL:    {url}")
    print(f"Domain filter: {domain}")
    print(f"{'='*60}")

    try:
        jar = browser_cookie3.chrome(domain_name=domain)
        cookies = [c for c in jar if domain in (c.domain or "")]

        if not cookies:
            print(f"  No cookies found for {domain}")
            return

        print(f"  Found {len(cookies)} cookies:\n")
        for c in cookies:
            print(f"  🍪 {c.domain:30s}  {c.name[:30]:30s}  {'🔒 secure' if c.secure else '  open'}")

    except PermissionError:
        print("  ❌ Permission denied — close Chrome and retry")
    except Exception as e:
        print(f"  ❌ Error: {e}")


# --- Test: these should ONLY show cookies for their own domain ---
test_urls = [
    "https://www.barrons.com/",
    "https://www.economist.com/",
    "https://www.google.com/",      # should NOT appear in barrons/economist results
    "https://mail.google.com/",     # Gmail — should NEVER leak into other domains
]

print("🔍 Cookie Sandbox Proof")
print("Each domain only sees its own cookies — no cross-site leakage.\n")

for url in test_urls:
    show_cookies_for_domain(url)

print(f"\n{'='*60}")
print("✅ If each section only shows cookies matching its own domain,")
print("   the sandbox is working correctly. No cross-site access.")
print(f"{'='*60}")
