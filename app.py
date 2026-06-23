from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import threading

# Import your scraping functions
from scraper import fetch_result  # assume you put your fetch_result function in scraper.py

app = Flask(__name__)
CORS(app)

# Endpoint to fetch results for multiple USNs
@app.route("/get_results", methods=["POST"])
def get_results():
    data = request.json
    college = data.get("college").upper()
    year = data.get("year")
    branch = data.get("branch").upper()
    low = int(data.get("low"))
    high = int(data.get("high"))
    sem = data.get("semester")

    results = []

    def run_scraper():
        for u in range(low, high + 1):
            usn = f"{college}{year}{branch}{str(u).zfill(3)}"
            res = fetch_result(usn, sem)  # Modify fetch_result to accept semester
            if res:
                results.append(res)

    thread = threading.Thread(target=run_scraper)
    thread.start()
    thread.join()

    return jsonify({"results": results})

if __name__ == "__main__":
    app.run(debug=True)
