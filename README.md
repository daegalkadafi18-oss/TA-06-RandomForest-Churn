# TA-06 – Random Forest: Prediksi Customer Churn

Dataset Kaggle: https://www.kaggle.com/datasets/blastchar/telco-customer-churn

## Isi folder
| Berkas | Fungsi |
|---|---|
| `TA06_RandomForest.ipynb` | Preprocessing, Skenario A (`n_estimators`), Skenario B (`max_depth`), Feature Importance, evaluasi, simpan model |
| `app.py`, `templates/index.html` | Deployment web (Flask) – form input pelanggan → probabilitas churn |
| `requirements.txt`, `Procfile` | Dependensi & perintah start untuk hosting (Render/Railway) |
| `data/` | Diisi otomatis (dataset diunduh saat notebook jalan) |
| `model/`, `figures/` | Dibuat otomatis oleh notebook |

## Cara menjalankan
1. `pip install -r requirements.txt jupyter`
2. Dataset diunduh otomatis oleh notebook saat dijalankan (kagglehub, cadangan salinan publik); tidak perlu unduh manual
3. Jalankan seluruh sel notebook (`jupyter notebook TA06_RandomForest.ipynb`) → folder `model/` dan `figures/` terisi
4. `python app.py` → buka http://127.0.0.1:5000
5. Deploy: push ke GitHub (sertakan `model/`), lalu buat Web Service di Render dengan start command `gunicorn app:app`
