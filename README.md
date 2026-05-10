# 🚚 Logistics Distance & Demand Forecasting: Prophet ML Pipeline

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Pandas](https://img.shields.io/badge/Pandas-Data_Manipulation-150458.svg)
![Prophet](https://img.shields.io/badge/Prophet-Time_Series-orange.svg)
![Gradio](https://img.shields.io/badge/Gradio-UI-FF7C00.svg)
![License](https://img.shields.io/badge/License-MIT-success.svg)

> **⚠️ Security & Confidentiality Notice:** This repository contains a conceptual, sanitized version of a production-level predictive pipeline. All confidential business logic, real pricing metrics, and proprietary company datasets have been removed. The pipeline utilizes dynamically generated dummy data to demonstrate architectural methodology and software engineering best practices.

## 📊 Business Problem
In large-scale logistics and distribution, resource allocation (trucks, fuel, drivers) relies heavily on accurate distance (kilometers) and demand forecasting. Traditional linear forecasting often fails to capture complex seasonality, holiday impacts, or sudden shifts in operational volume, leading to misallocation of budget and assets. 

This project demonstrates an automated Machine Learning pipeline and UI designed to forecast transportation distances, allowing operations teams to preemptively adjust routing and budget strategies.

## ⚙️ Architecture & Methodology
The pipeline takes raw historical travel data, processes it for anomalies, and utilizes **Facebook Prophet** for robust time-series forecasting, all wrapped in a professional **Gradio** web interface.

1. **Data Ingestion & Cleaning:** Automated ETL steps using `pandas` to handle missing values, smooth outliers using IQR (Interquartile Range), and aggregate daily distance metrics.
2. **Feature Engineering:** Incorporation of custom seasonality (weekly, monthly) and national holidays that impact distribution volumes.
3. **Predictive Modeling:** Model training using `prophet` with dynamic regressors (monthly units).
4. **Professional Output:** Generates interactive multi-panel visualizations (Matplotlib) and exports highly formatted Excel reports (using `openpyxl`) for downstream Business Intelligence tools.

## 🚀 Key Features

* **High-Precision Time Series:** Configured to handle complex seasonality and holidays.
* **Automated Outlier Detection:** Identifies and removes statistical anomalies before model training to prevent skewed forecasts.
* **Confidence Intervals & Risk Assessment:** Provides upper and lower bounds for predictions, allowing stakeholders to assess statistical risk.
* **Production-Ready UI:** A dark-mode, responsive web interface built with Gradio for seamless user interaction without requiring coding knowledge.

## 🛠️ How to Run (Demo Version)

1. Clone this repository:
   ```bash
   git clone [https://github.com/your-username/Logistics-Distance-Forecasting-Prophet.git](https://github.com/your-username/Logistics-Distance-Forecasting-Prophet.git)
