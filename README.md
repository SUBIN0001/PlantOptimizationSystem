# 🏭 Ashok Leyland — Plant Location Decision System

> A full-stack Multi-Criteria Decision Making (MCDM) web application for optimal industrial plant location selection, powered by AHP + Entropy + TOPSIS and Monte Carlo Simulation.

---

## 📌 Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [Project Structure](#project-structure)
- [Getting Started](#getting-started)
  - [Backend Setup](#backend-setup)
  - [Frontend Setup](#frontend-setup)
- [Dataset / Excel Input](#dataset--excel-input)
- [Usage](#usage)
- [API Endpoints](#api-endpoints)
- [Algorithms Used](#algorithms-used)

---

## Overview

This system helps industrial planners evaluate and rank potential plant locations across India based on six key criteria:

| Criterion | Type | Description |
|---|---|---|
| **Vendor Base** | Benefit | Auto-component supplier density within 200 km |
| **Manpower Availability** | Benefit | Engineering college & ITI graduate pool |
| **CAPEX** | Cost | Total estimated project capital expenditure |
| **Govt / Norms** | Benefit | State incentives, subsidies, and approvals ease |
| **Logistics Cost** | Cost | Annual freight and transport cost |
| **Economies of Scale** | Benefit | Cluster maturity and market demand index |

The backend computes AHP weights (via pairwise comparison matrix), Entropy weights, and a **hybrid score** blended at a configurable alpha (α). Final ranking is done using **TOPSIS**. Sensitivity analysis is performed via **Monte Carlo simulation (1000+ iterations)**.

---

## Features

- 📤 **Upload Excel datasets** — supports both detailed sub-attribute format and simplified pre-scored format
- ⚖️ **Pairwise AHP comparison** — interactive matrix with consistency ratio check
- 🔬 **Hybrid AHP + Entropy weighting**
- 🏆 **TOPSIS ranking** with feasibility constraint filtering
- 🎲 **Monte Carlo simulation** for sensitivity and rank stability analysis
- 📊 **Styled Excel export** with 5 sheets: Raw Data, TOPSIS Results, AHP Weights, Entropy Weights, Monte Carlo
- 🌍 **Region & State filter** support
- 🎨 Animated React UI with Ashok Leyland branding

---

## Tech Stack

### Backend
| Library | Version | Purpose |
|---|---|---|
| FastAPI | 0.115.0 | REST API framework |
| Uvicorn | 0.30.6 | ASGI server |
| Pandas | 2.2.2 | Excel parsing & data processing |
| NumPy | 1.26.4 | Numerical computation |
| Scikit-learn | 1.5.1 | ML utilities |
| Openpyxl | 3.1.5 | Excel read/write |
| XlsxWriter | latest | Styled Excel export |
| python-multipart | 0.0.9 | File upload handling |

### Frontend
| Library | Version | Purpose |
|---|---|---|
| React | 19.x | UI framework |
| Vite | 8.x | Build tool & dev server |
| XLSX | 0.18.5 | Client-side spreadsheet utility |
| Vanilla CSS | — | Custom styling |

---

## Project Structure

```
curr_proj/
├── backend/
│   ├── app.py               # FastAPI main application (all MCDM logic)
│   ├── excel_generator.py   # Styled Excel export utilities
│   └── requirements.txt     # Python dependencies
│
└── frontend/
    └── src/
        └── ashok-leyland-ui/
            ├── src/
            │   ├── App.jsx          # Main React app
            │   └── ...              # Components, pages, assets
            ├── public/              # Static assets
            ├── index.html
            ├── package.json
            └── vite.config.js
```

---

## Getting Started

### Prerequisites
- **Python** 3.10+
- **Node.js** 18+ and **npm**

---

### Backend Setup

```bash
# Navigate to the backend folder
cd backend

# (Recommended) Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux

# Install dependencies
pip install -r requirements.txt

# Start the backend server
uvicorn app:app --reload --port 8000
```

The API will be available at: `http://localhost:8000`  
Interactive API docs: `http://localhost:8000/docs`

---

### Frontend Setup

```bash
# Navigate to the frontend app folder
cd frontend/src/ashok-leyland-ui

# Install Node dependencies
npm install

# Start the development server
npm run dev
```

The app will be available at: `http://localhost:5173`

> **Note:** Make sure the backend is running before using upload or analysis features.

```gitignore
# .gitignore — add these lines
*.xlsx
*.xls
*.csv
data/
```

---

### Supported Excel Formats

The backend auto-detects the format on upload.

#### Format 1 — Detailed Sub-Attribute Sheet (`finaleyy.xlsx`)
A wide table where each row is a location and columns represent raw sub-attributes. The backend computes the 6 core scores automatically.

**Required columns include (sample):**
- `Location`, `Region`, `State`
- `No. of ACMA Member Units (State, approx.)`
- `Tier-1 Auto Vendors within 200 km (nos.)`
- `Estimated Total Project CAPEX (₹ Cr)*`
- `Capital Subsidy (% of Fixed Assets)`
- `Annual Logistics Cost (₹ Cr/yr, est.)**`
- `Auto Industry Cluster Maturity`
- *(and ~35 more sub-attribute columns)*

#### Format 2 — Simplified Pre-Scored Sheet
A compact table with one aggregated score per dimension per location.

**Required columns:**
| Column | Example Value | Description |
|---|---|---|
| `Location` | Pune | Plant site name |
| `Region` | West | Geographic region |
| `State` | Maharashtra | Indian state |
| `Vendor Base` | `8.5` | Score 0–10 (benefit) |
| `Manpower Availability` | `7.2` | Score 0–10 (benefit) |
| `CAPEX` | `₹ 62 Cr` or `62` | Capital expenditure (₹ Crores) |
| `Govt Norms` | `High` or `7.0` | Score or text rating |
| `Logistics Cost` | `Low` or `3.0` | Score or text rating |
| `Economies of Scale` | `Medium` | Text rating or score 0–10 |

Text ratings accepted: `Very High`, `High`, `Medium`, `Low`, `Very Low`

---

## Usage

1. **Launch** both backend and frontend servers.
2. **Upload** your Excel file on the Upload page.
3. **Review** the parsed location table.
4. **Set constraints** (e.g., CAPEX ≤ 80 Cr, min Vendor Base score = 6).
5. **Fill in** the AHP pairwise comparison matrix for criteria weights.
6. **Run Analysis** to get ranked results via TOPSIS.
7. **Run Monte Carlo** simulation to check rank stability.
8. **Export** results to a styled multi-sheet Excel report.

---

## API Endpoints

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/upload-excel` | Upload and parse Excel dataset |
| `POST` | `/api/analyze` | Run AHP + Entropy + TOPSIS analysis |
| `POST` | `/api/monte-carlo` | Run Monte Carlo simulation |
| `POST` | `/api/export-excel` | Generate styled Excel report |
| `GET` | `/health` | Health check |

---

## Algorithms Used

### 1. AHP (Analytic Hierarchy Process)
- User provides an **n×n pairwise comparison matrix** for the 6 criteria.
- Weights are derived via eigenvector normalization.
- **Consistency Ratio (CR)** is computed; CR < 0.1 is acceptable.

### 2. Entropy Weighting
- Objective weights derived from the **information entropy** of each criterion across all locations.
- High variation in a criterion → higher entropy weight.

### 3. Hybrid Weight
```
W_hybrid = α × W_AHP + (1 − α) × W_Entropy
```
where α is configurable (default 0.5).

### 4. TOPSIS
- Normalized decision matrix weighted by hybrid weights.
- Ideal best (A⁺) and ideal worst (A⁻) solutions computed per criterion type (benefit/cost).
- Closeness coefficient `C = S⁻ / (S⁺ + S⁻)` determines rank.

### 5. Monte Carlo Simulation
- **1000 iterations** with ±15% weight perturbation and ±5% value perturbation.
- Outputs: average rank, rank standard deviation, 95% confidence interval, rank probability distribution.

---

## License

This project is developed as an academic & analytical tool for **Ashok Leyland** plant location decision-making.  
© 2026 — All rights reserved.
