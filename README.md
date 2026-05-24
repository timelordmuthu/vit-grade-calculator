# 🎓 VIT FFCS Grade Calculator v2

An unofficial, enhanced grade predictor for VIT University students built with Streamlit. Supports all three FFCS course types — Pure Theory, Pure Lab, and Embedded (Theory + Lab) — and follows FFCS Regulations v4.0.

---

## ✨ Features

- **Three course type modes** — Pure Theory, Pure Lab, and Embedded (Theory + Lab)
- **Relative & Absolute grading** — auto-switches based on course type; manually overridable
- **Sigma estimation** — estimates standard deviation from class mean and topper's mark, or accept a manual value
- **CAM calculator** — computes Continuous Assessment Marks from CAT-I, CAT-II, and three Digital Assignments
- **FAT target table** — shows the FAT score needed to achieve each grade (S through F)
- **Theory FAT Simulator** — interactive slider to simulate different FAT outcomes and see predicted grade in real time
- **Embedded credit-weight ratio** — automatically derived from LTPC credits, with an optional manual override
- **Grade Band Chart** — visual bar chart showing where your score falls relative to all grade cutoffs
- **Persistent Excel storage** — save, load, and manage multiple subjects locally in `vit_grades_data.xlsx`
- **Import / Export Excel** — download your data, edit it freely, and re-import to sync
- **Auto-fill** — selecting a previously saved subject pre-fills all fields

---

## 📋 Grading Rules Applied (FFCS v4.0)

| Course Type | CAM (60%) | FAT (40%) | Pass Rule | Grading |
|---|---|---|---|---|
| **Pure Theory** | CAT-I×0.3 + CAT-II×0.3 + DA(3×10) | Theory FAT /100 × 0.4 | FAT ≥ 40/100 | Relative (Absolute if ≤10 students) |
| **Pure Lab** | Lab CA avg /100 × 0.6 | Lab FAT /100 × 0.4 | Lab total ≥ 50/100 | Absolute always |
| **Embedded** | Theory CAM /60 + Lab CA × 0.6 | Theory FAT × 0.4 + Lab FAT × 0.4 | Lab ≥ 50 AND Theory FAT ≥ 40 | `(Theory×(L+T) + Lab×P/2) / (L+T+P/2)` |

**Relative grade bands:**
- S ≥ mean + 1.5σ (min 90 for Pure Theory; capped at 80 for Embedded)
- A ≥ mean + 0.5σ | B ≥ mean − 0.5σ | C ≥ mean − 1σ | D ≥ mean − 1.5σ | E ≥ mean − 2σ

**Sigma estimation:** σ ≈ (topper − class mean) / 2.5

---

## 🚀 Getting Started

### Prerequisites

- Python 3.8+
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/your-username/vit-grade-calculator.git
cd vit-grade-calculator

# Install dependencies
pip install streamlit pandas openpyxl
```

### Running the App

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.

---

## 🖥️ How to Use

1. **Step ① — Subject & Course Type**
   Enter your subject name and select the course type (Pure Theory / Pure Lab / Embedded).

2. **Step ② — Class Statistics**
   Enter the class average and topper's mark to get accurate relative grade band predictions. Leave blank to use the default σ = 10.

3. **Step ③ — Enter Your Marks**
   Fill in your CAT-I, CAT-II, Digital Assignments, and Lab scores. Use the FAT Simulator slider to explore different FAT outcomes.

4. **Step ④ — Save & Manage**
   Save subjects to Excel, download for offline editing, re-import changes, or delete entries.

---

## 📁 Project Structure

```
vit-grade-calculator/
│
├── app.py                  # Main Streamlit application
├── vit_grades_data.xlsx    # Auto-generated on first save (local storage)
└── README.md
```

> `vit_grades_data.xlsx` is created automatically when you save your first subject. You can add it to `.gitignore` if you don't want to commit personal grade data.

---

## ⚠️ Disclaimer

This is an **unofficial** grade predictor and is **not affiliated with VIT University**. Predictions are estimates based on publicly available FFCS regulations. Always verify your grades with your faculty or official VIT portals.

---

## 🤝 Contributing

Pull requests are welcome! If you find a bug or want to suggest an improvement, feel free to open an issue.

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).
