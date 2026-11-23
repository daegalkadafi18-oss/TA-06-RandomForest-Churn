"""Bangun Laporan TA-06 (format jurnal COMPUTING, 2 kolom) dari template jurnal.

Pemakaian:
    python make_report.py                 # memakai model/results.json + figures/ bila ada
    python make_report.py --placeholder   # paksa mode kerangka (semua angka berupa [isi: ...])
Hasil: Laporan_TA-06_Jurnal_301240005_DaegalFauzanKadafi.docx
Prasyarat: pip install python-docx ; file template_jurnal.docx berada di folder yang sama.
Bagian yang masih berwarna kuning = wajib dilengkapi manual (URL GitHub/Deploy/YouTube, screenshot web).
"""
import copy
import json
import os
import re
import sys

import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.text.paragraph import Paragraph

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template_jurnal.docx")
OUT = os.path.join(HERE, "Laporan_TA-06_Jurnal_301240005_DaegalFauzanKadafi.docx")
FORCE_PH = "--placeholder" in sys.argv

# ------------------------------------------------------------------ data
R = None
dfA = dfB = None
rp = os.path.join(HERE, "model", "results.json")
if not FORCE_PH and os.path.exists(rp):
    R = json.load(open(rp))
    dfA = pd.read_csv(os.path.join(HERE, "figures", "hasil_skenario_a.csv"))
    dfB = pd.read_csv(os.path.join(HERE, "figures", "hasil_skenario_b.csv"))
HAS = R is not None


def idn(x, d=2):
    s = f"{x:,.{d}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x, d=2): return idn(x * 100, d)
def ph(label): return f"[[isi: {label}]]"
def V(fn, label):  # nilai bila ada hasil, selain itu placeholder
    return fn() if HAS else ph(label)


# ------------------------------------------------------------------ template
d = Document(TEMPLATE)
body = d.element.body
els = list(body)
sect = body.find(qn("w:sectPr"))
proto = {
    "head": copy.deepcopy(els[8]), "text": copy.deepcopy(els[9]), "eq": copy.deepcopy(els[17]),
    "cap_t": copy.deepcopy(els[20]), "gap": copy.deepcopy(els[22]), "sub": copy.deepcopy(els[29]),
    "pic": copy.deepcopy(els[31]), "cap_f": copy.deepcopy(els[32]), "refh": copy.deepcopy(els[52]),
    "ref": copy.deepcopy(els[53]),
}
# paragraf 0-6 = judul, penulis, abstrak (diedit), 7 = pemisah seksi (dipertahankan)
head_paras = [Paragraph(e, d._body) for e in els[:7]]
for e in els[8:]:
    if e is not sect:
        body.remove(e)

TOKEN = re.compile(r"(\*[^*]+\*|\[\[[^\]]+\]\])")


def fill(par, text, bold=None, italic=None):
    for r in list(par.runs):
        r._r.getparent().remove(r._r)
    for part in TOKEN.split(text):
        if not part:
            continue
        if part.startswith("[["):
            r = par.add_run("[" + part[2:-2] + "]")
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            r = par.add_run(part[1:-1]); r.italic = True
        else:
            r = par.add_run(part)
        if bold is not None and part and not part.startswith("[["):
            r.bold = bold
        if italic and not part.startswith("[["):
            r.italic = True
    return par


def add(kind, text="", **kw):
    el = copy.deepcopy(proto[kind])
    for ch in list(el):
        if ch.tag != qn("w:pPr"):
            el.remove(ch)
    sect.addprevious(el)
    p = Paragraph(el, d._body)
    if text:
        fill(p, text, **kw)
    return p


def H(t): add("head", t, bold=True)
def S(t): add("sub", t, bold=True, italic=True)
def T(t): add("text", t)


def eq(formula, n):
    p = add("eq")
    p.add_run("\t" + formula + "\t(" + str(n) + ")")


def figure(fname, caption, label):
    path = os.path.join(HERE, "figures", fname)
    p = add("pic")
    if HAS and os.path.exists(path):
        p.add_run().add_picture(path, width=Inches(3.35))
    else:
        r = p.add_run(f"[Gambar: {label} – dihasilkan notebook, simpan di figures/{fname}]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    add("cap_f", caption)


def table(caption, header, rows, widths):
    add("cap_t", caption)
    tbl = OxmlElement("w:tbl")
    pr = OxmlElement("w:tblPr")
    w = OxmlElement("w:tblW"); w.set(qn("w:w"), str(sum(widths))); w.set(qn("w:type"), "dxa"); pr.append(w)
    tbl.append(pr)
    grid = OxmlElement("w:tblGrid")
    for x in widths:
        g = OxmlElement("w:gridCol"); g.set(qn("w:w"), str(x)); grid.append(g)
    tbl.append(grid)

    def cell(text, wd, first, top, bottom, bold, keep):
        tc = OxmlElement("w:tc")
        tcpr = OxmlElement("w:tcPr")
        tw = OxmlElement("w:tcW"); tw.set(qn("w:w"), str(wd)); tw.set(qn("w:type"), "dxa"); tcpr.append(tw)
        b = OxmlElement("w:tcBorders")
        for side, on in (("top", top), ("left", False), ("bottom", bottom), ("right", False)):
            e = OxmlElement(f"w:{side}")
            e.set(qn("w:val"), "single" if on else "none"); e.set(qn("w:sz"), "6" if on else "0")
            e.set(qn("w:space"), "0"); e.set(qn("w:color"), "000000" if on else "FFFFFF"); b.append(e)
        tcpr.append(b)
        m = OxmlElement("w:tcMar")
        for side, v in (("top", 30), ("left", 60), ("bottom", 30), ("right", 60)):
            e = OxmlElement(f"w:{side}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
        tcpr.append(m)
        tc.append(tcpr)
        pe = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        if keep:
            ppr.append(OxmlElement("w:keepNext"))
        if not first:
            jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "center"); ppr.append(jc)
        pe.append(ppr)
        tc.append(pe)
        par = Paragraph(pe, d._body)
        fill(par, text)
        for r in par.runs:
            r.font.size = __import__("docx").shared.Pt(9)
            if bold:
                r.bold = True
        return tc

    allrows = [(header, True)] + [(r, False) for r in rows]
    for i, (row, is_head) in enumerate(allrows):
        tr = OxmlElement("w:tr")
        last = i == len(allrows) - 1
        for j, (txt, wd) in enumerate(zip(row, widths)):
            tr.append(cell(str(txt), wd, j == 0, is_head, is_head or last, is_head, not last))
        tbl.append(tr)
    sect.addprevious(tbl)
    add("gap")


# ------------------------------------------------------------------ nilai turunan
rep = R["report"] if HAS else None
best_n = V(lambda: str(R["best_n"]), "n_estimators terbaik")
best_d = V(lambda: "None (tanpa batas)" if R["best_depth"] is None else str(R["best_depth"]), "max_depth terbaik")
test_acc = V(lambda: pct(R["test_acc"]) + "%", "akurasi testing")
train_acc = V(lambda: pct(R["train_acc"]) + "%", "akurasi training")
rec1 = V(lambda: pct(rep["Churn"]["recall"]) + "%", "recall churn")
pre1 = V(lambda: pct(rep["Churn"]["precision"]) + "%", "precision churn")
f11 = V(lambda: pct(rep["Churn"]["f1-score"]) + "%", "F1 churn")
n_rows = V(lambda: idn(R["n_rows"], 0), "jumlah baris")
churn_share = V(lambda: pct(R["n_churn"] / R["n_rows"], 1) + "%", "persentase churn")

# ------------------------------------------------------------------ judul & abstrak
fill(head_paras[0], "OPTIMASI HYPERPARAMETER RANDOM FOREST UNTUK PREDIKSI CHURN PELANGGAN TELEKOMUNIKASI "
     "DAN IMPLEMENTASINYA PADA APLIKASI WEB", bold=True)
fill(head_paras[1], "Daegal Fauzan Kadafi", bold=True)
fill(head_paras[2], "Program Studi Teknik Informatika, Universitas Bale Bandung")

abs_id = ("*ABSTRAK:* Pelanggan yang berhenti berlangganan (*churn*) menimbulkan kerugian bagi perusahaan telekomunikasi "
          "sehingga kemampuan memprediksinya sangat dibutuhkan. Penelitian ini bertujuan membangun model *Random Forest* "
          "yang akurat, stabil, dan dapat dijelaskan untuk memprediksi *churn* serta menyajikannya dalam aplikasi web. "
          f"Data yang digunakan ialah *Telco Customer Churn* dari Kaggle sebanyak {n_rows} pelanggan dengan proporsi "
          f"*churn* {churn_share}. Metode meliputi pembersihan data, *one-hot encoding*, pembagian data 80:20 berstrata, "
          "serta dua skenario eksperimen: pengaruh jumlah pohon (*n_estimators*) dan kedalaman pohon (*max_depth*). "
          f"Konfigurasi terbaik ialah n_estimators = {best_n} dan max_depth = {best_d} dengan akurasi testing "
          f"{test_acc} (akurasi training {train_acc}), serta recall kelas *churn* {rec1} dan F1-score {f11}. "
          "Analisis *feature importance* mengidentifikasi variabel kunci yang berpengaruh terhadap keputusan pelanggan. "
          "Model disimpan dan diintegrasikan ke aplikasi Flask yang di-*deploy* ke *cloud*. Simpulannya, Random Forest "
          "dengan konfigurasi tersebut layak dipakai sebagai alat bantu retensi pelanggan.")
abs_en = ("*ABSTRACT:* Customer churn causes financial losses for telecommunication companies, so predicting it is valuable. "
          "This study aims to build an accurate, stable, and explainable Random Forest model to predict churn and to deliver it "
          f"in a web application. The data is the Telco Customer Churn dataset from Kaggle with {n_rows} customers and a churn "
          f"share of {churn_share}. The method includes data cleaning, one-hot encoding, a stratified 80:20 split, and two "
          "experimental scenarios: the effect of the number of trees (n_estimators) and tree depth (max_depth). The best "
          f"configuration is n_estimators = {best_n} and max_depth = {best_d}, giving a testing accuracy of {test_acc} "
          f"(training accuracy {train_acc}), a churn recall of {rec1}, and an F1-score of {f11}. Feature importance analysis "
          "identifies the key variables driving customer decisions. The model is saved and integrated into a Flask application "
          "deployed to the cloud. In conclusion, the tuned Random Forest is suitable as a customer retention support tool.")
fill(head_paras[3], abs_id, bold=True)
# bold hanya untuk label: set ulang -> label tebal, isi tidak
for par, text in ((head_paras[3], abs_id), (head_paras[5], abs_en)):
    fill(par, text)
    for r in par.runs:
        r.bold = False
        r.italic = r.italic
    # label (ABSTRAK:/ABSTRACT:) tebal-miring
    first = par.runs[0]
    first.bold = True
fill(head_paras[4], "*Kata kunci:* Random Forest, churn pelanggan, hyperparameter, feature importance, Flask")
fill(head_paras[6], "*Keywords:* Random Forest, customer churn, hyperparameter, feature importance, Flask")
for par in (head_paras[3], head_paras[4], head_paras[5], head_paras[6]):
    for r in par.runs:
        if not r.font.highlight_color:
            r.italic = True

# ------------------------------------------------------------------ PENDAHULUAN
H("PENDAHULUAN")
T("Pada industri telekomunikasi, biaya memperoleh pelanggan baru jauh lebih tinggi daripada mempertahankan pelanggan lama, "
  "sehingga pelanggan yang berhenti berlangganan (*churn*) menjadi masalah bisnis yang serius (Verbeke dkk., 2012). "
  "Prediksi *churn* membantu perusahaan menentukan pelanggan yang perlu diberi penawaran retensi lebih dahulu. "
  "Masalah ini berbentuk klasifikasi biner dengan data tabular yang memuat atribut numerik dan kategorikal, serta "
  "proporsi pelanggan *churn* yang umumnya lebih kecil daripada yang bertahan (Vafeiadis dkk., 2015).")
T("*Random Forest* merupakan metode *ensemble* yang membangun banyak pohon keputusan dan menggabungkan prediksinya melalui "
  "pemungutan suara mayoritas (Breiman, 2001). Setiap pohon dilatih pada sampel *bootstrap* dari data latih, yaitu "
  "pengambilan acak dengan pengembalian (*bagging*) (Breiman, 1996), dan pada setiap percabangan hanya sebagian fitur "
  "acak yang dipertimbangkan sehingga pohon-pohon menjadi beragam dan galat variansnya turun (Hastie dkk., 2009). Karena itu "
  "*Random Forest* lebih tahan terhadap *overfitting* daripada satu pohon keputusan, tahan terhadap *noise*, dan tidak "
  "memerlukan normalisasi fitur.")
T("Kinerja *Random Forest* dipengaruhi oleh *hyperparameter*, terutama jumlah pohon (*n_estimators*) dan kedalaman maksimum "
  "pohon (*max_depth*). Penambahan pohon pada umumnya menstabilkan prediksi hingga titik jenuh, sedangkan pohon yang terlalu "
  "dalam dapat menghafal data latih (Probst dkk., 2019). Selain akurat, model perlu dapat dijelaskan. *Random Forest* "
  "menyediakan *feature importance* berdasarkan penurunan impuritas rata-rata (*mean decrease in impurity*), meskipun ukuran "
  "ini dapat condong pada fitur dengan banyak nilai unik (Strobl dkk., 2007), sehingga hasilnya perlu ditafsirkan bersama "
  "pengetahuan domain.")
T("Penelitian ini tidak membandingkan algoritma lain, melainkan mengoptimalkan satu model *Random Forest* agar akurat dan tidak "
  "*overfitting*. Rumusan masalahnya ialah bagaimana pengaruh *n_estimators* dan *max_depth* terhadap kinerja model, fitur "
  "apa yang paling berpengaruh terhadap *churn*, dan bagaimana mengintegrasikan model terbaik ke aplikasi web. Tujuannya ialah "
  "menemukan konfigurasi terbaik melalui eksperimen terkontrol, mengidentifikasi variabel kunci, dan menyajikan model "
  "dalam aplikasi web yang dapat diakses publik.")

# ------------------------------------------------------------------ METODA
H("METODA")
T("Penelitian ini bersifat eksperimental kuantitatif. Data berasal dari *Telco Customer Churn* di Kaggle (blastchar) yang memuat "
  f"{n_rows} pelanggan dengan atribut demografi, layanan yang dilanggan, jenis kontrak, metode pembayaran, lama berlangganan "
  "(*tenure*), serta biaya bulanan dan total. Variabel target ialah *Churn* (Yes = berhenti, No = bertahan). Percobaan dilakukan "
  "pada November 2026 menggunakan Python, Pandas, scikit-learn (Pedregosa dkk., 2011), Matplotlib/Seaborn, dan Flask.")
T("Pra-pemrosesan meliputi penghapusan kolom *customerID*, konversi *TotalCharges* menjadi numerik (nilai kosong pada pelanggan "
  "dengan *tenure* nol diisi 0), serta pengecekan ketidakseimbangan kelas pada target. Seluruh atribut kategorikal diubah menjadi "
  "numerik dengan *one-hot encoding* dan target diubah menjadi 0/1. Data dibagi menjadi 80% data latih dan 20% data uji secara "
  "berstrata dengan *random_state* 42 agar proporsi *churn* terjaga pada kedua bagian"
  + (f" ({idn(R['n_train'], 0)} data latih dan {idn(R['n_test'], 0)} data uji)." if HAS else " ([[isi: jumlah data latih dan uji]])."))
T("Dilakukan dua skenario eksperimen (Tabel 1) dengan *random_state* tetap. Pada Skenario A, jumlah pohon divariasikan dan "
  "dipilih nilai terbaik berdasarkan akurasi testing, dengan memilih jumlah yang lebih kecil bila selisihnya tidak lebih dari "
  "0,2 poin persentase agar lebih efisien. Pada Skenario B, dengan *n_estimators* terbaik tersebut, kedalaman pohon divariasikan "
  "dan akurasi training dibandingkan dengan akurasi testing; *overfitting* dinilai dari selisih (*gap*) keduanya. Konfigurasi akhir "
  "ialah kedalaman dengan akurasi testing tertinggi dan *gap* terkecil.")
table("Tabel 1 Rancangan eksperimen", ["Skenario", "Parameter yang divariasikan", "Tetap"],
      [["A", "n_estimators: 10, 50, 100, 200, 500", "max_depth = None"],
       ["B", "max_depth: None, 3, 5, 10, 20", "n_estimators terbaik A"]], [900, 2700, 1300])
T("Kinerja dievaluasi dengan akurasi, *precision*, *recall*, dan *F1-score* pada Persamaan 1 sampai 3, dengan TP, FP, dan FN "
  "berturut-turut adalah jumlah prediksi positif benar, positif salah, dan negatif salah pada kelas *churn*. Pada kasus ini "
  "*recall* kelas *churn* paling penting karena pelanggan *churn* yang tidak terdeteksi tidak akan mendapat penawaran retensi.")
eq("Precision = TP / (TP + FP)", 1)
eq("Recall = TP / (TP + FN)", 2)
eq("F1 = 2 · Precision · Recall / (Precision + Recall)", 3)
T("Model terbaik diambil kemudian *feature importance*-nya diekstrak dari atribut *feature_importances_* dan 10 fitur teratas "
  "divisualisasikan. Model disimpan dengan Joblib bersama metadata fitur, lalu dimuat oleh aplikasi web Flask yang menerima "
  "input data pelanggan, menerapkan *one-hot encoding* yang sama, dan menampilkan probabilitas *churn* beserta tingkat risikonya. "
  "Aplikasi di-*deploy* ke layanan *cloud*.")

# ------------------------------------------------------------------ HASIL
H("HASIL DAN PEMBAHASAN")
S("Eksplorasi Data")
if HAS:
    imb = R["n_stay"] / R["n_churn"]
    T(f"Data berjumlah {n_rows} pelanggan dengan {R['n_features_raw']} atribut dan menghasilkan {R['n_features_encoded']} fitur "
      f"setelah *encoding*. Pelanggan *churn* berjumlah {idn(R['n_churn'], 0)} ({churn_share}) dan yang bertahan "
      f"{idn(R['n_stay'], 0)}, atau rasio sekitar {idn(imb, 1)} : 1 (Gambar 1). Ketidakseimbangan ini membuat akurasi saja "
      "kurang memadai, sehingga *recall* dan *F1-score* kelas *churn* ikut diperhatikan.")
else:
    T(f"Data berjumlah {n_rows} pelanggan. Pelanggan *churn* berjumlah {ph('jumlah churn')} dan yang bertahan {ph('jumlah bertahan')} "
      "(Gambar 1). [[isi: analisis ketidakseimbangan kelas dan dampaknya pada pemilihan metrik.]]")
figure("class_distribution.png", "Gambar 1 Distribusi kelas target (Churn)", "distribusi kelas")

S("Skenario A: Pengaruh Jumlah Pohon")
if HAS:
    a = dfA
    rowsA = [[int(r.n_estimators), pct(r.train_acc), pct(r.test_acc)] for r in a.itertuples()]
else:
    rowsA = [[n, ph("angka"), ph("angka")] for n in (10, 50, 100, 200, 500)]
table("Tabel 2 Hasil Skenario A (akurasi, %)", ["n_estimators", "Training", "Testing"], rowsA, [1500, 1600, 1600])
if HAS:
    lo, hi = a.loc[a.test_acc.idxmin()], a.loc[a.test_acc.idxmax()]
    gain = (hi.test_acc - a.test_acc.iloc[0]) * 100
    plateau = int(a[a.test_acc >= hi.test_acc - 0.005].n_estimators.min())
    t_a = (f"Berdasarkan Tabel 2 dan Gambar 2, akurasi testing berada pada rentang {pct(lo.test_acc)}% (n = {int(lo.n_estimators)}) "
           f"sampai {pct(hi.test_acc)}% (n = {int(hi.n_estimators)}). ")
    if gain >= 1.0:
        t_a += (f"Kenaikan dari 10 pohon ke nilai terbaik sebesar {idn(gain)} poin persentase, tetapi kurva mulai mendatar sejak "
                f"sekitar {plateau} pohon sehingga penambahan berikutnya memberi manfaat yang kecil. ")
    else:
        t_a += (f"Selisih antara 10 pohon dan nilai terbaik hanya {idn(gain)} poin persentase, sehingga menambah pohon "
                "terus-menerus tidak meningkatkan akurasi secara signifikan. ")
    t_a += ("Hal ini sesuai sifat *Random Forest*: penambahan pohon menurunkan varians prediksi hingga titik jenuh dan tidak "
            f"menimbulkan *overfitting* tambahan, tetapi menambah waktu pelatihan dan memori. Dipilih n_estimators = {best_n}.")
    T(t_a)
else:
    T("[[isi: analisis Skenario A – apakah menambah pohon terus-menerus meningkatkan akurasi secara signifikan, dan pada jumlah pohon "
      "berapa kurva mulai mendatar.]]")
figure("skenario_a.png", "Gambar 2 Jumlah pohon terhadap akurasi testing", "grafik Skenario A")

S("Skenario B: Pengaruh Kedalaman Pohon")
if HAS:
    rowsB = [[r.max_depth, pct(r.train_acc), pct(r.test_acc), pct(r.gap)] for r in dfB.itertuples()]
else:
    rowsB = [[dd, ph("angka"), ph("angka"), ph("angka")] for dd in ("None", 3, 5, 10, 20)]
table("Tabel 3 Hasil Skenario B (akurasi, %)", ["max_depth", "Training", "Testing", "Gap"], rowsB, [1200, 1200, 1200, 1100])
if HAS:
    order = ["3", "5", "10", "20", "None"]
    b = dfB.copy(); b["max_depth"] = b["max_depth"].astype(str); b = b.set_index("max_depth").loc[order]
    over = [k for k in order if b.loc[k, "gap"] > 0.10]
    t_b = (f"Pada Tabel 3 dan Gambar 3, akurasi training pada kedalaman 3 sebesar {pct(b.loc['3', 'train_acc'])}% dan terus naik "
           f"hingga {pct(b.loc['None', 'train_acc'])}% pada kedalaman tak terbatas, sedangkan akurasi testing berada pada "
           f"{pct(b.test_acc.min())}%–{pct(b.test_acc.max())}%. ")
    if over:
        t_b += (f"Selisih training–testing melebihi 10 poin persentase mulai kedalaman {over[0]}, sehingga *overfitting* mulai "
                "terlihat di kedalaman tersebut: model menghafal data latih tanpa peningkatan akurasi uji yang sepadan. ")
    else:
        t_b += "Selisih training–testing tidak pernah melebihi 10 poin persentase sehingga tidak tampak *overfitting* yang nyata. "
    t_b += (f"Berdasarkan kriteria akurasi testing tertinggi dengan *gap* terkecil, dipilih max_depth = {best_d}. Akurasi "
            "training mendekati 100% pada pohon dalam wajar terjadi pada *Random Forest* karena setiap pohon melihat sampel "
            "*bootstrap*, sehingga yang menentukan ialah akurasi testing dan bukan akurasi training semata.")
    T(t_b)
else:
    T("[[isi: analisis Skenario B – pada kedalaman berapa selisih training–testing mulai melebar (overfitting), dan kedalaman yang dipilih.]]")
figure("skenario_b_bar.png", "Gambar 3 Akurasi training dan testing pada berbagai kedalaman", "grafik Skenario B")

S("Feature Importance")
BIZ = {
    "tenure": ("*tenure*", "pelanggan baru jauh lebih rentan pergi, sedangkan pelanggan lama sudah terikat dan puas",
               "program *onboarding* dan penawaran khusus pada beberapa bulan pertama"),
    "Contract": ("jenis kontrak", "kontrak bulanan tidak mengikat sehingga mudah dihentikan, sedangkan kontrak satu atau dua tahun "
                 "menahan pelanggan", "insentif untuk beralih dari kontrak bulanan ke kontrak tahunan"),
    "MonthlyCharges": ("biaya bulanan", "tagihan yang tinggi meningkatkan sensitivitas harga dan dorongan mencari penyedia lain",
                       "paket yang lebih sesuai atau diskon bagi pelanggan bertagihan tinggi"),
    "TotalCharges": ("total tagihan", "mencerminkan gabungan lama berlangganan dan besar tagihan, sehingga berkaitan erat dengan "
                     "*tenure*", "penghargaan loyalitas bagi pelanggan dengan total tagihan besar"),
    "InternetService": ("jenis layanan internet", "layanan serat optik biasanya lebih mahal dan pelanggannya lebih sensitif pada "
                        "kualitas", "pemantauan kualitas jaringan dan harga layanan tersebut"),
    "PaymentMethod": ("metode pembayaran", "pembayaran manual seperti *electronic check* lebih mudah diputus dibandingkan debit otomatis",
                      "ajakan beralih ke pembayaran otomatis"),
    "TechSupport": ("dukungan teknis", "pelanggan tanpa dukungan teknis lebih mudah kecewa saat menghadapi masalah",
                    "bundel layanan dukungan teknis"),
    "OnlineSecurity": ("keamanan daring", "layanan tambahan meningkatkan keterikatan pelanggan", "paket gabungan layanan keamanan"),
}
if HAS:
    top = R["top10"]
    names = [t["Feature"] for t in top]
    base = lambda f: next((k for k in BIZ if f.startswith(k)), None)
    seen, expl, acts = [], [], []
    for f in names:
        k = base(f)
        if k and k not in seen:
            seen.append(k)
    for k in seen[:5]:
        lbl, why, act = BIZ[k]
        expl.append(f"{lbl} ({why})")
        acts.append(act)
    T(f"Gambar 4 menunjukkan 10 fitur teratas pada model terbaik. Fitur paling penting ialah {names[0]} "
      f"({idn(top[0]['Importance'], 3)}), diikuti {names[1]} ({idn(top[1]['Importance'], 3)}) dan {names[2]} "
      f"({idn(top[2]['Importance'], 3)}). "
      + ("Secara domain, hasil ini masuk akal: " + "; ".join(expl) + ". " if expl else "")
      + "Fitur-fitur ini sejalan dengan temuan umum riset *churn* telekomunikasi bahwa lama berlangganan, jenis kontrak, dan "
        "besar tagihan merupakan penentu utama. Perlu dicatat bahwa *feature importance* menunjukkan kekuatan hubungan, bukan "
        "sebab-akibat, dan dapat condong pada fitur numerik berskala kontinu (Strobl dkk., 2007).")
else:
    T("[[isi: interpretasi 10 fitur teratas – jelaskan logika di balik fitur penting (mis. tenure, Contract, MonthlyCharges) dan apakah "
      "sesuai teori domain churn.]]")
figure("feature_importance.png", "Gambar 4 Sepuluh fitur terpenting model Random Forest", "feature importance")

S("Evaluasi Model Terbaik")
if HAS:
    rows = []
    for k, nm in (("Tidak Churn", "Tidak Churn"), ("Churn", "Churn")):
        rows.append([nm, pct(rep[k]["precision"]), pct(rep[k]["recall"]), pct(rep[k]["f1-score"]), int(rep[k]["support"])])
    rows.append(["Akurasi", "", "", pct(R["test_acc"]), int(rep["macro avg"]["support"])])
    rows.append(["Macro avg", pct(rep["macro avg"]["precision"]), pct(rep["macro avg"]["recall"]), pct(rep["macro avg"]["f1-score"]),
                 int(rep["macro avg"]["support"])])
else:
    rows = [[nm, ph("angka"), ph("angka"), ph("angka"), ph("angka")] for nm in ("Tidak Churn", "Churn", "Akurasi", "Macro avg")]
table(f"Tabel 4 Classification report model terbaik (n_estimators = {best_n}, max_depth = {best_d}; %)",
      ["Kelas", "Precision", "Recall", "F1", "Support"], rows, [1300, 900, 900, 900, 900])
if HAS:
    cm = R["cm"]
    fn, fp = cm[1][0], cm[0][1]
    T(f"Model terbaik mencapai akurasi testing {test_acc} dan akurasi training {train_acc}. Untuk kelas *churn*, *precision* "
      f"{pre1}, *recall* {rec1}, dan *F1-score* {f11} (Tabel 4). Pada {idn(R['n_test'], 0)} data uji terdapat {fn} pelanggan "
      f"*churn* yang tidak terdeteksi (*false negative*) dan {fp} pelanggan bertahan yang salah ditandai berisiko (*false positive*). "
      "Karena kelas *churn* lebih sedikit, kinerjanya lebih rendah daripada kelas mayoritas. Bagi bisnis, *false negative* "
      "lebih merugikan daripada *false positive*, sehingga ke depan ambang keputusan dapat diturunkan atau diterapkan "
      "pembobotan kelas (*class_weight*) untuk menaikkan *recall*.")
    T("Implikasi bisnis: " + ("; ".join(acts) + ". " if acts else "")
      + "Pelanggan dengan probabilitas *churn* tinggi pada aplikasi web dapat langsung diprioritaskan untuk tindakan retensi.")
else:
    T("[[isi: analisis metrik kelas churn (precision/recall/F1), jumlah false negative & false positive, dan implikasi bisnis yang dapat ditindaklanjuti.]]")

S("Implementasi Aplikasi Web")
T("Aplikasi web dibangun dengan Flask. Pengguna mengisi data pelanggan (atribut numerik dan pilihan kategori) pada formulir, "
  "aplikasi mengubahnya ke fitur yang sama dengan data latih, lalu menampilkan probabilitas *churn*, tingkat risiko "
  "(rendah, sedang, tinggi), dan saran tindakan, serta daftar 10 fitur terpenting model (Gambar 5). Model dan metadata dimuat "
  "sekali saat aplikasi dijalankan, kemudian aplikasi di-*deploy* ke layanan *cloud* sehingga dapat diakses publik.")
shot = os.path.join(HERE, "figures", "web_app.png")
p = add("pic")
if os.path.exists(shot):
    p.add_run().add_picture(shot, width=Inches(3.35))
else:
    r = p.add_run("[Gambar: screenshot aplikasi web – simpan sebagai figures/web_app.png lalu jalankan ulang make_report.py]")
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
add("cap_f", "Gambar 5 Tampilan aplikasi prediksi churn pelanggan")

# ------------------------------------------------------------------ SIMPULAN
H("SIMPULAN DAN SARAN")
if HAS:
    T(f"Model *Random Forest* untuk prediksi *churn* berhasil dibangun dan dioptimalkan. Konfigurasi terbaik adalah n_estimators = "
      f"{best_n} dan max_depth = {best_d} dengan akurasi testing {test_acc}, recall kelas *churn* {rec1}, dan F1-score {f11}. "
      "Jumlah pohon yang lebih banyak menstabilkan hasil hingga titik jenuh, sedangkan kedalaman pohon menentukan jarak antara "
      "akurasi training dan testing. Fitur paling berpengaruh ialah "
      + ", ".join(names[:3]) + ". Model dan aplikasi web telah terintegrasi dan di-*deploy*.")
else:
    T("[[isi: ringkasan konfigurasi terbaik (n_estimators & max_depth), metrik utama kelas churn, fitur terpenting, dan status deployment.]]")
T("Saran untuk penelitian selanjutnya ialah menangani ketidakseimbangan kelas (misalnya *class_weight* atau *SMOTE*), menyetel "
  "*hyperparameter* lain seperti *min_samples_leaf* dan *max_features* dengan *cross-validation*, serta menyesuaikan ambang "
  "keputusan untuk menaikkan *recall* kelas *churn*. Pemilihan konfigurasi pada penelitian ini memakai data uji yang sama "
  "dengan evaluasi akhir, sehingga penelitian lanjutan sebaiknya memakai data validasi terpisah.")

# ------------------------------------------------------------------ PUSTAKA
H("PUSTAKA ACUAN")
add("refh", "Buku", bold=True)
add("ref", "Hastie, T., Tibshirani, R., and Friedman, J. (2009). The Elements of Statistical Learning (2nd ed.). New York, USA: Springer.")
add("refh", "Jurnal/Prosiding/Disertasi/Tesis/Skripsi", bold=True)
for r_ in [
    "Breiman, L. (1996). Bagging Predictors. Machine Learning, 24(2), 123-140.",
    "Breiman, L. (2001). Random Forests. Machine Learning, 45(1), 5-32.",
    "Pedregosa, F., dkk. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12, 2825-2830.",
    "Probst, P., Wright, M. N., and Boulesteix, A.-L. (2019). Hyperparameters and Tuning Strategies for Random Forest. WIREs Data Mining and Knowledge Discovery, 9(3), e1301.",
    "Strobl, C., Boulesteix, A.-L., Zeileis, A., and Hothorn, T. (2007). Bias in Random Forest Variable Importance Measures: Illustrations, Sources and a Solution. BMC Bioinformatics, 8, 25.",
    "Vafeiadis, T., Diamantaras, K. I., Sarigiannidis, G., and Chatzisavvas, K. C. (2015). A Comparison of Machine Learning Techniques for Customer Churn Prediction. Simulation Modelling Practice and Theory, 55, 1-9.",
    "Verbeke, W., Dejaeger, K., Martens, D., Hur, J., and Baesens, B. (2012). New Insights into Churn Prediction in the Telecommunication Sector: A Profit Driven Data Mining Approach. European Journal of Operational Research, 218(1), 211-229.",
]:
    add("ref", r_)
add("refh", "Lain-lain", bold=True)
add("ref", "Kaggle. Telco Customer Churn (blastchar/telco-customer-churn). https://www.kaggle.com/datasets/blastchar/telco-customer-churn")
add("ref", "Github : [[isi URL GitHub]]")
add("ref", "Deploy : [[isi URL deployment]]")
add("ref", "Yt : [[isi URL YouTube]]")

d.save(OUT)
print("OK ->", OUT, "| mode:", "HASIL NYATA" if HAS else "KERANGKA (placeholder)")
