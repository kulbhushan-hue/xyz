import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import ptf_scraper as ps

PAGE = """<a href="/p2.html">PTF list 2</a><table><tr><th>PTF</th><th>FMID</th><th>APAR</th><th>Date</th><th>Description</th></tr>
<tr><td>UJ98216</td><td>HDBD11K</td><td>PH12345</td><td>2026-03-02</td><td>Support z/OS 3.2</td></tr></table>"""
PAGES = {"http://x/a.html": PAGE, "http://x/p2.html": "<p>Apply BQU1234 released 14 Mar 2026</p>",
         "http://x/login.html": "<input type=password>Sign in"}

def fake_get(s, url, retries=2):
    return (PAGES[url], None) if url in PAGES else (None, "Access denied")

def test_scrape(monkeypatch=None):
    ps._get = fake_get
    r = ps.scrape_url("http://x/a.html", "Db2 13.1", "IBM")
    assert [x["PTF ID"] for x in r] == ["UJ98216", "BQU1234"] and r[0]["FMID"] == "HDBD11K"
    assert ps.scrape_url("http://x/login.html", "P", "BMC")[0]["Detail Extraction Status"] == "Access denied"
    assert ps.scrape_url("http://x/none.html", "P", "BMC")[0]["Detail Extraction Status"] == "Access denied"

if __name__ == "__main__":
    test_scrape(); print("PASS")
