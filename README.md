# IT Ticket Analyzer - Professional Edition

A professional, enterprise-ready ticket tracking and analysis system built with Python and Streamlit. This tool helps you manage and analyze IT support tickets across multiple workflow stages with enhanced filtering, user statistics, and AI-powered reply assistance.

## ✨ Key Features

### 🎯 **Enhanced Ticket Tracking**
- **Multi-stage workflow tracking**: Work in Progress → Under Development → Awaiting User Info
- **Smart search functionality**: Find tickets instantly with comprehensive status overview
- **Professional dashboard**: Clean, modern interface suitable for customer demonstrations

### 📊 **Advanced Analytics & Filtering**
- **User statistics with animated counters**: Real-time statistics showing ticket distribution per owner
- **Comprehensive workload analysis**: Identify overloaded team members and bottlenecks
- **Smart filtering**: Filter by owner across all categories with live statistics updates
- **Workload alerts**: Automatic warnings for high-priority situations and overloaded users

### 🤖 **AI-Powered Assistant** 
- **Intelligent reply drafting**: AI-generated customer follow-up messages
- **Context-aware responses**: Tailored messages based on ticket description and context
- **OpenAI integration**: Powered by GPT models for professional communication

### 🎨 **Professional Design**
- **Enterprise-ready interface**: Clean, professional styling suitable for business environments
- **Responsive design**: Works seamlessly across desktop and mobile devices  
- **Modern UI components**: Professional cards, animated transitions, and intuitive navigation
- **Customer-ready appearance**: Polished interface ready for client presentations

## 🚀 What's New in Professional Edition

### Enhanced User Experience
- **Animated statistics**: Smooth counter animations when viewing user statistics
- **Color-coded metrics**: Visual indicators for different ticket categories and priorities
- **Professional styling**: Clean white background with modern card-based layout
- **Smart alerts**: Context-aware warnings and workload management insights

### Advanced Filtering & Analytics
- **Global owner filtering**: View complete statistics for any team member across all categories
- **Tab-level filtering**: Dedicated filters within each workflow stage
- **Percentage breakdowns**: Understand workload distribution at a glance
- **Real-time updates**: Statistics update instantly when filters are applied

### Professional Features
- **Workload management**: Automatic detection of overloaded team members
- **Priority indicators**: Visual priority coding in related ticket displays
- **Action alerts**: Intelligent warnings for tickets requiring immediate attention
- **Enterprise styling**: Professional appearance suitable for customer demonstrations

## 📋 Prerequisites

- **Python 3.9+** - [Download here](https://www.python.org/downloads/)
- **Git** - For cloning and version control
- **Excel files** - Your ticket exports in .xlsx format

## 🛠️ Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/SandeeepKrish/IT-Ticket-Analyzer.git
cd IT-Ticket-Analyzer
```

### 2. Create Virtual Environment
```bash
python -m venv .venv
```

**Activate the environment:**
- **Windows PowerShell**: `.venv\Scripts\Activate.ps1`
- **Windows Command Prompt**: `.venv\Scripts\activate.bat`
- **macOS/Linux**: `source .venv/bin/activate`

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Environment (Optional)
For AI features, copy `.env.example` to `.env` and add your OpenAI API key:
```bash
OPENAI_API_KEY=your_openai_api_key_here
```

### 5. Launch the Application
```bash
streamlit run app.py
```

The application will automatically open in your browser at `http://localhost:8501`

## 📱 How to Use

### Step 1: Upload Ticket Data
In the sidebar, upload your three Excel exports:
- **Work in Progress** (.xlsx)
- **Under Development** (.xlsx)  
- **Awaiting User Info** (.xlsx)

### Step 2: Explore Your Data

#### 🔍 **Search Individual Tickets**
- Enter any ticket number to see its complete status
- View detailed ticket information with professional card layout
- See all related tickets for the same owner

#### 📊 **Analyze User Workloads**  
- Use the "Filter by Owner" dropdown to select any team member
- View animated statistics showing their ticket distribution
- Get smart alerts about workload and urgent items

#### 📈 **Browse by Category**
- Navigate through the three main tabs
- Filter within each category by owner
- View real-time statistics and data tables

### Step 3: AI-Powered Assistance (Optional)
- Add your OpenAI API key in the sidebar
- Click "✍️ Draft reply" next to any ticket awaiting user info
- Get professionally crafted follow-up messages

## 🎯 Key Use Cases

### For Managers
- **Team workload monitoring**: See who's overloaded and needs help
- **Bottleneck identification**: Find tickets stuck waiting for responses
- **Performance analytics**: Track ticket distribution and completion patterns

### For Support Teams  
- **Quick ticket lookup**: Find any ticket's status instantly
- **Context awareness**: See all related tickets for better customer service
- **AI assistance**: Get help drafting professional customer communications

### For Customers & Stakeholders
- **Professional interface**: Clean, enterprise-ready appearance for demos
- **Real-time insights**: Live statistics and filtering capabilities
- **Comprehensive reporting**: Complete visibility into ticket workflows

## 📊 Excel File Requirements

Your Excel exports should include these column headers (exact names):
```
Ticket Number, Ticket Creator, Subject, Customer Name, 
Ticket Department, Status, Ticket Owner, Ticket Priority, 
Ticket Aging, Start Date, Required Date, 
Estimate Resolution Date, Description
```

- **All three files** (Work in Progress, Under Development, Awaiting User Info) should have the same column structure
- **Extra columns** are fine and will be preserved
- **Column order** doesn't matter as long as headers match

## 🛡️ Security & Privacy

- **Local processing**: All data stays on your machine
- **No data transmission**: Files are processed locally, not sent to external servers
- **API key protection**: OpenAI keys are stored securely and never logged
- **Environment variables**: Sensitive configuration protected via .env files

## 🔧 Technical Architecture

### Built With
- **Python 3.9+** - Core runtime
- **Streamlit** - Web application framework
- **Pandas + OpenPyXL** - Excel file processing and data analysis
- **OpenAI API** - AI-powered reply generation (optional)

### Project Structure
```
IT-Ticket-Analyzer/
├── app.py              # Main Streamlit application
├── ai.py               # AI reply generation module
├── requirements.txt    # Python dependencies
├── README.md          # Documentation
├── .gitignore         # Git ignore rules
└── .env.example       # Environment configuration template
```

## 🚀 Deployment Options

### Local Development
Perfect for personal use and team testing. Simply run `streamlit run app.py`.

### Enterprise Deployment
For production use, consider:
- **Streamlit Cloud**: Easy cloud deployment
- **Docker containers**: Containerized deployment
- **Internal servers**: Deploy on company infrastructure

## 🤝 Contributing

We welcome contributions! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature-name`
3. **Make your changes** with tests
4. **Submit a pull request** with a clear description

### Development Guidelines
- Follow Python PEP 8 style guidelines
- Add comments for complex logic
- Test with various Excel file formats
- Ensure responsive design works on mobile

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support & Troubleshooting

### Common Issues

**Excel files not loading?**
- Ensure files have .xlsx extension
- Check that all required columns are present
- Verify file isn't password protected

**AI features not working?**
- Verify OpenAI API key is valid
- Check your OpenAI account has available credits
- Ensure .env file is properly configured

**Application won't start?**  
- Confirm Python 3.9+ is installed
- Verify virtual environment is activated
- Run `pip install -r requirements.txt` again

### Getting Help
- **GitHub Issues**: Report bugs or request features
- **Documentation**: Check this README for detailed information
- **Community**: Join discussions in the repository

## 🎉 Acknowledgments

- Built with [Streamlit](https://streamlit.io) for the web interface
- Powered by [OpenAI](https://openai.com) for AI assistance
- Uses [Pandas](https://pandas.pydata.org) for data processing

---

**Made with ❤️ for IT support teams everywhere**
