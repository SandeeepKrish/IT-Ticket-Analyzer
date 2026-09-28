# IT Ticket Analyzer - Enterprise Edition 🎫

An enterprise-grade IT ticket flow tracking, analytics, and automated follow-up platform built with **Python**, **Streamlit**, and **OpenAI**. 

This platform streamlines support operations across 8 workflow stages, eliminates bottlenecked requests with intelligent creator email matching, automated background reminders, and features a data-aware AI chatbot.

---

## ✨ Key Features & Capabilities

### 1. 🔄 Multi-Stage Ticket Flow Tracking
- **Complete Lifecycle Coverage**: Tracks tickets across **8 distinct categories**:
  - `Master Ticket List`, `Open`, `Work In Progress (WIP)`, `Under Development`, `Awaiting User Info`, `On Hold`, `Pending`, and `Closed`.
- **Cross-Queue Search**: Search any ticket number or creator to instantly view its live status and uncover every other open ticket associated with that owner.
- **Workflow Process Visualizer**: Interactive diagram outlining ticket progression from creation to resolution.

### 2. 📁 Smart Batch / Folder Upload with Auto-Detection
- **Single-Click Ingestion**: Upload all Excel files (`.xlsx`) at once in a single folder upload.
- **Fuzzy Auto-Detection**: Automatically identifies and categorizes files using filename matching (e.g. `master`, `wip`, `dev`, `awaiting`, `hold`, `open`, `pending`, `closed`, and `mail_to`).
- **Validation & Summaries**: Live status indicators confirming all workflow files and directories are detected and validated before processing.

### 3. ✉️ Automated Mail Services & Follow-Up Reminders
- **Mail Directory Integration (`mail_to.xlsx`)**: Dynamically resolves creator names to corporate email addresses using exact, lowercase, and partial token matching.
- **1-Click "Send via Outlook"**: Encoded `mailto:` URL links pre-filled with ticket number, subject, aging, and response guidance.
- **Bulk Background SMTP Dispatch**: Send hundreds of personalized follow-up emails directly through your corporate mail server (Office 365, Gmail, custom SMTP) with live progress tracking.
- **Queue Follow-up Cards**: Visual cards displaying creator email tags, aging counters, and expandable draft previews.

### 4. 🤖 AI-Powered Assistant & Chatbot
- **Data-Aware Chatbot**: Query your uploaded tickets using natural language. The chatbot understands ticket counts, departments, bottlenecks, and aging data.
- **Smart Reply Drafting**: Generates context-aware follow-up messages tailored to ticket descriptions using OpenAI GPT models.

### 5. 📊 Advanced Analytics & Date Filtering
- **Dynamic Slicers**: Filter by Owner, Year, and Month dynamically across all tabs.
- **Workload & Bottleneck Metrics**: Visual KPI counters, aging analyses, and workload distribution breakdowns.

---

## 🏗️ Project Architecture & Structure

The codebase is built on a clean, modular architecture:

```text
ticket-flow-tracker-py/
│
├── app.py              # Main dashboard application, page routing, and Streamlit UI
├── mail_service.py     # ✉️ Mail directory parsing, email matching, mailto generation & SMTP dispatch
├── ai.py               # 🤖 OpenAI integration, AI reply drafting, chatbot & folder auto-detection
├── footer.py           # 🏷️ Branded enterprise footer component
├── style.css           # 🎨 Custom styling, glassmorphism, responsive cards & status badges
│
├── requirements.txt    # Python dependencies
├── .env.example        # Environment variable configuration template
├── .gitignore          # Version control ignore rules
└── README.md           # Project documentation
```

### Module Responsibilities:
- **[app.py](file:///y:/Common/SandeepERPandIT/SANDDEP%20TECH/ticket-flow-tracker-py/app.py)**: Coordinates UI layout, sidebar uploads, filtering engines, and tab navigation.
- **[mail_service.py](file:///y:/Common/SandeepERPandIT/SANDDEP%20TECH/ticket-flow-tracker-py/mail_service.py)**: Encapsulates all email logic (`build_email_directory`, `find_creator_email`, `generate_reminder_email`, `generate_mailto_url`, `send_smtp_email`).
- **[ai.py](file:///y:/Common/SandeepERPandIT/SANDDEP%20TECH/ticket-flow-tracker-py/ai.py)**: Manages OpenAI client connections, prompt engineering, dataset context serialization, and filename fuzzy detection.
- **[style.css](file:///y:/Common/SandeepERPandIT/SANDDEP%20TECH/ticket-flow-tracker-py/style.css)**: Lightweight CSS overriding Streamlit defaults with custom gradients, badges, and card components.

---

## 🛠️ Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/SandeeepKrish/IT-Ticket-Analyzer.git
cd IT-Ticket-Analyzer
```

### 2. Set Up Virtual Environment
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables
Copy `.env.example` to `.env` (or create `.env`):
```ini
# AI Features (Optional)
OPENAI_API_KEY=your_openai_api_key_here

# SMTP Automated Email Settings (Optional)
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_PASSWORD=your_app_password_here
```

### 5. Run the Application
```bash
streamlit run app.py
```
The application will launch at `http://localhost:8501`.

---

## 📋 Data & File Requirements

### Ticket Workflow Files (`.xlsx`)
Export files (`Master`, `WIP`, `Under Development`, `Awaiting User Info`, `Hold`, `Open`, `Pending`, `Closed`) should include the following standard columns:
- `Ticket Number`
- `Ticket Creator`
- `Subject`
- `Ticket Department`
- `Status`
- `Ticket Owner`
- `Ticket Priority`
- `Ticket Aging`
- `Start Date`
- `Customer Name` (optional)
- `Required Date` / `Estimate Resolution Date` (optional)

### Mail Directory File (`mail_to.xlsx`)
To enable automated email matching, provide an Excel file containing:
- **Name Column**: `Ticket Creator`, `Name`, `User`, `Employee Name`, etc.
- **Email Column**: `Email`, `Email ID`, `Mail`, `Mail To`, etc.

---

## 🔒 Security & Privacy

- **Local Execution**: All Excel files and ticket information are processed in-memory on your local machine.
- **Credential Protection**: Passwords and API keys are loaded via `.env` and kept out of version control.
- **Zero Third-Party Storage**: Ticket data is never uploaded to external databases.

---

## 📄 License
This project is licensed under the [MIT License](LICENSE).
