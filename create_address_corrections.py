from flask import Flask, jsonify, request
import requests
import threading
import webbrowser

app = Flask(__name__)

PAGE = """<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Census Address Tester</title>
<style>
body{font-family:Arial,sans-serif;max-width:760px;margin:40px auto;padding:0 20px}
input{width:100%;box-sizing:border-box;padding:12px;font-size:16px}
button{margin-top:10px;padding:10px 18px;font-size:16px}
#status{margin-top:18px;font-weight:bold}
pre{background:#111;color:#eee;padding:12px;white-space:pre-wrap}
</style>
</head>
<body>
<h1>Census Address Tester</h1>
<form id="form">
  <input id="address" placeholder="631 N Reed St, Sisters, OR 97759" required>
  <button type="submit">Submit</button>
</form>
<div id="status"></div>
<pre id="output"></pre>

<script>
document.getElementById("form").addEventListener("submit", async (event) => {
  event.preventDefault();

  const address = document.getElementById("address").value.trim();
  const status = document.getElementById("status");
  const output = document.getElementById("output");

  status.textContent = "Checking...";
  output.textContent = "";

  try {
    const response = await fetch("/geocode?address=" + encodeURIComponent(address));
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Request failed");
    }

    const matches = data?.result?.addressMatches || [];

    status.textContent = matches.length ? "MATCH" : "NO MATCH";
    output.textContent = JSON.stringify(data, null, 2);
  } catch (error) {
    status.textContent = "ERROR: " + error.message;
  }
});
</script>
</body>
</html>
"""


@app.route("/")
def home():
    return PAGE


@app.route("/geocode")
def geocode():
    address = request.args.get("address", "").strip()

    if not address:
        return jsonify({"error": "Address is required"}), 400

    try:
        response = requests.get(
            "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress",
            params={
                "address": address,
                "benchmark": "Public_AR_Current",
                "format": "json",
            },
            timeout=30,
        )
        response.raise_for_status()
        return jsonify(response.json())
    except requests.RequestException as exc:
        return jsonify({"error": str(exc)}), 502


if __name__ == "__main__":
    url = "http://127.0.0.1:8765/"
    print(f"Open: {url}")
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=8765, debug=False)