from flask import Flask, request, render_template, send_file, redirect, url_for, flash
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd, time, re
import uuid, os

app = Flask(__name__)
app.secret_key = os.urandom(24)

def extract_title(html):
    soup = BeautifulSoup(html, "lxml")
    tag = soup.find("span", class_="pageheader")
    if tag:
        return tag.get_text(strip=True)
    if soup.title:
        m = re.search(r'(5\.\d+\.\d+[^ ]*)|Day0[^|\-]*', soup.title.get_text())
        if m: return m.group(0)
        return soup.title.get_text(strip=True)
    m = re.search(r'<span[^>]*class=["\']?pageheader["\']?[^>]*>([^<]+)</span>', html)
    return m.group(1).strip() if m else "Unknown Title"

def extract_data(html):
    soup = BeautifulSoup(html, "lxml")
    cfg = [span.get_text(strip=True) for span in soup.select("#confsDIV .link-normal")] or ["<none>"]
    rows = []
    for r in soup.select("tr[style*='vertical-align:top']"):
        cols = r.find_all("td")
        if len(cols) < 11: continue
        rows.append({
            "Module": cols[0].get_text(strip=True),
            "Passed": cols[3].get_text(strip=True).split()[0],
            "Failed": cols[5].get_text(strip=True).split()[0],
            "Quality": cols[9].get_text(strip=True),
            "Defects": cols[10].get_text(strip=True) if len(cols)>10 else ""
        })
    return cfg, rows

def scrape_tims(tims_id):
    driver = webdriver.Chrome()
    driver.get("https://tims.cisco.com")
    input("🔐 Please log in to Cisco TIMS manually, then press ENTER...")
    driver.get(f"https://tims.cisco.com/warp.cmd?ent=Tnr{tims_id}")
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    time.sleep(2)
    html = driver.page_source
    driver.quit()
    title = extract_title(html)
    configs, rows = extract_data(html)
    df = pd.DataFrame(rows)
    fname = f"tims_{tims_id}_{uuid.uuid4().hex[:8]}.xlsx"
    df.to_excel(fname, index=False)
    return title, configs, fname

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        tid = request.form.get("tims_id", "").strip()
        if not tid:
            flash("Enter a TIMS ID")
            return redirect(url_for("index"))
        title, cfgs, fname = scrape_tims(tid)
        return render_template("results.html", title=title, configs=cfgs, download_file=url_for("download", fname=fname))
    return render_template("index.html")

@app.route("/download/<fname>")
def download(fname):
    return send_file(fname, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
