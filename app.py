"""Aplikasi web (Flask) – Prediksi Churn Pelanggan dengan Random Forest.

Jalankan lokal :  python app.py   -> buka http://127.0.0.1:5000
Produksi       :  gunicorn app:app   (mis. di Render / Railway / Hugging Face Spaces)
Prasyarat      :  jalankan notebook dulu agar folder model/ (rf_churn.joblib & meta.json) terbentuk.
"""
import json
import os

import joblib
import pandas as pd
from flask import Flask, render_template, request

BASE = os.path.dirname(os.path.abspath(__file__))
MODEL_DIR = os.environ.get("MODEL_DIR", os.path.join(BASE, "model"))

app = Flask(__name__)
model = joblib.load(os.path.join(MODEL_DIR, "rf_churn.joblib"))
with open(os.path.join(MODEL_DIR, "meta.json")) as f:
    meta = json.load(f)

# Pelanggan contoh agar form mudah dicoba
DEFAULT_CAT = {"Contract": "Month-to-month", "InternetService": "Fiber optic"}


def build_features(form):
    """Ubah input form -> DataFrame berfitur sama persis dengan data latih (One-Hot + urutan kolom)."""
    row = {}
    for c in meta["num_cols"]:
        row[c] = float(form.get(c, meta["num_stats"][c]["median"]))
    for c in meta["cat_cols"]:
        row[c] = form.get(c, meta["cat_options"][c][0])
    raw = pd.DataFrame([row])
    enc = pd.get_dummies(raw, columns=meta["cat_cols"], dtype=int)
    return enc.reindex(columns=meta["feature_columns"], fill_value=0)


def risk_label(p):
    if p >= 0.6:
        return "Tinggi", "high", "Segera tawarkan retensi: diskon, upgrade kontrak jangka panjang, atau follow-up layanan."
    if p >= 0.35:
        return "Sedang", "mid", "Pantau pelanggan dan tawarkan paket/kontrak yang lebih sesuai."
    return "Rendah", "low", "Pelanggan relatif aman; pertahankan kualitas layanan."


@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    values = {}
    if request.method == "POST":
        values = request.form.to_dict()
        X = build_features(request.form)
        prob = float(model.predict_proba(X)[0][1])
        label, css, advice = risk_label(prob)
        result = {"prob": round(prob * 100, 1), "label": label, "css": css, "advice": advice,
                  "churn": prob >= 0.5}
    return render_template("index.html", meta=meta, result=result, values=values, default_cat=DEFAULT_CAT)


@app.route("/health")
def health():
    return {"status": "ok", "n_features": len(meta["feature_columns"])}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
