Energy Solar AI

An AI-powered solar load calculator that automates electricity bill analysis and generates solar recommendation reports.

Users can upload electricity bills in PDF or image format, and the system automatically extracts billing details using Groq LLMs and generates a filled Excel report with solar sizing, savings, ROI, and cost calculations.

The project is built using:

Streamlit (Frontend)
Groq API + LLaMA Models (AI Extraction)
openpyxl (Excel Automation)
PyMuPDF (PDF Processing)
🚀 Features
Upload electricity bill (PDF/Image)
AI-based bill data extraction
Automatic solar system recommendation
ROI & savings calculation
Download ready-to-use Excel report
Custom solar assumptions from sidebar
⚙️ Run Locally
pip install -r requirements.txt
streamlit run app.py

Set your Groq API key:
Windows CMD
set GROQ_API_KEY=your_api_key

🌐 Live Demo
[Energy Solar AI Demo](https://energy-solar-ai.onrender.com/)

📂 GitHub Repository

Energy Solar AI GitHub Repository

🔮 Future Improvements
Better bill-format parsing
WhatsApp integration
CRM automation
Cloud report storage
