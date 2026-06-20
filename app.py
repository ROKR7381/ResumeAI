import streamlit as st
import json
import base64
import os
from dotenv import load_dotenv
from schemas import ResumeData, PersonalInfo, WorkExperience, Education, SkillGroup, Project, Certification, Language
from agent import refine_resume_with_agent, generate_resume_suggestions
from exporter import export_to_excel, render_html_resume

# Load environment variables from .env file (force override on reload)
load_dotenv(override=True)


# Page Configuration
st.set_page_config(
    page_title="AI Resume Builder & Co-Pilot",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Rich Aesthetics)
st.markdown("""
<style>
    /* Theme overrides */
    .stApp {
        background-color: #0F172A; /* Slate 900 */
        color: #F8FAFC; /* Slate 50 */
    }
    
    /* Input fields styling */
    .stTextInput>div>div>input, .stTextArea>div>div>textarea {
        background-color: #1E293B !important; /* Slate 800 */
        color: #F8FAFC !important;
        border: 1px solid #334155 !important; /* Slate 700 */
    }
    .stTextInput>div>div>input:focus, .stTextArea>div>div>textarea:focus {
        border-color: #38BDF8 !important; /* Sky 400 */
        box-shadow: 0 0 0 1px #38BDF8 !important;
    }
    
    /* Header card */
    .header-card {
        background: linear-gradient(135deg, #1E293B, #0F172A);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
    }
    .header-card h1 {
        color: #38BDF8;
        font-weight: 800;
        margin: 0 0 10px 0;
    }
    .header-card p {
        color: #94A3B8;
        margin: 0;
        font-size: 16px;
    }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #1E293B !important;
        border-right: 1px solid #334155;
    }
    
    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #0F172A;
        padding: 10px 0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 6px;
        color: #94A3B8;
        font-weight: 600;
        padding: 10px 20px;
        transition: all 0.3s;
    }
    .stTabs [aria-selected="true"] {
        background-color: #38BDF8 !important;
        color: #0F172A !important;
        border-color: #38BDF8 !important;
    }
    
    /* Accordion / Expander Card styling */
    .stExpander {
        background-color: #1E293B !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        margin-bottom: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to get base64 encoded link for new tab open
def get_html_download_link(html_code, filename="resume.html", link_text="Download HTML"):
    b64 = base64.b64encode(html_code.encode()).decode()
    return f'<a href="data:text/html;base64,{b64}" download="{filename}" class="download-btn" style="text-decoration:none; background-color:#38BDF8; color:#0F172A; padding:10px 20px; border-radius:6px; font-weight:bold; display:inline-block; margin-top:10px;">{link_text}</a>'

def get_html_print_link(html_code, link_text="🖨️ Print / Save to PDF"):
    b64 = base64.b64encode(html_code.encode()).decode()
    # Opens a new tab and triggers window.print() once loaded
    js_print = f"""
    <html>
    <head>
        <script>
            window.onload = function() {{
                window.print();
            }}
        </script>
    </head>
    <body style="margin:0;padding:0;">
        <iframe src="data:text/html;base64,{b64}" style="width:100vw; height:100vh; border:none; margin:0; padding:0;"></iframe>
    </body>
    </html>
    """
    b64_print = base64.b64encode(js_print.encode()).decode()
    return f'<a href="data:text/html;base64,{b64_print}" target="_blank" style="text-decoration:none; background-color:#10B981; color:white; padding:10px 20px; border-radius:6px; font-weight:bold; display:inline-block; margin-top:10px;">{link_text}</a>'

# Initialize Session State for Resume Data
if "resume_data" not in st.session_state:
    st.session_state.resume_data = {
        "personal_info": {
            "full_name": "",
            "professional_title": "",
            "email": "",
            "phone": "",
            "location": "",
            "linkedin": "",
            "github": "",
            "website": "",
            "birth_info": "",
            "nationality": "",
            "hobbies": "",
            "photo_base64": None,
            "power_statement": ""
        },
        "summary": "",
        "experience": [],
        "education": [],
        "skills": [],
        "projects": [],
        "certifications": [],
        "languages": [],
        "accomplishments": []
    }

# Sample Vague Data for Testing (Populated with Roshan's docx resume)
sample_data = {
    "personal_info": {
        "full_name": "Roshan Kumar",
        "professional_title": "Data Scientist",
        "email": "roshan.kr.hit@gmail.com",
        "phone": "6206194628",
        "location": "KOLKATA, INDIA",
        "linkedin": "http://www.linkedin.com/in/roshan7381",
        "github": "https://github.com/ROKR7381",
        "website": "https://sites.google.com/view/roshandatascientistportfolio/home",
        "birth_info": "05/03/1984, Bokaro",
        "nationality": "Indian",
        "hobbies": "Web_site Development, Learning and upgrading my skills",
        "photo_base64": None,
        "power_statement": "Leveraging Machine Learning & GenAI to solve complex financial and industrial challenges."
    },
    "summary": "Results-driven Data Scientist with extensive experience in machine learning, predictive modeling, and generative AI. Currently working on a major government project for UCO Bank, focusing on probability of default prediction, data-driven credit offers, and customer 360 views. Proven track record of developing RAG-based chatbots and deploying AI automation solutions to drive business value.",
    "experience": [
        {
            "company": "MOL IT",
            "position": "Senior Software Developer",
            "location": "Remote",
            "start_date": "Jul 2025",
            "end_date": "Present",
            "description": [
                "Developing and enhancing AI-based automation tools for shipping data analytics.",
                "Maintaining and improving a chatbot system that interacts with structured data to answer all shipping-related queries efficiently.",
                "Developed a chatbot application that accepts structured shipment data as input and provides accurate, context-aware answers to shipping queries such as container details, ETA, and shipment tracking.",
                "Utilized Python, Pandas, LLMs, Streamlit, RAG Architecture, and Azure."
            ]
        },
        {
            "company": "INSPIRA ENTERPRISE INDIA LTD",
            "position": "Data Scientist",
            "location": "Kolkata, India",
            "start_date": "Apr 2023",
            "end_date": "Jul 2025",
            "description": [
                "Currently working on UCO BANK Government Project.",
                "Designed and built machine learning models for Probability of Default (PD) prediction for various loan categories (Home, Car, Agriculture, MSME) to assist in risk assessment.",
                "Engineered systems for data-driven credit offers and single view of customer (Customer 360).",
                "Developed and deployed a Retrieval-Augmented Generation (RAG) chatbot application for shipping data analysis on Microsoft Azure, combining NLP with domain-specific shipping dataset retrieval.",
                "Addressed multicollinearity using Variance Inflation Factor (VIF) and balanced imbalanced datasets using SMOTE.",
                "Trained and evaluated various classification models; achieved 86% accuracy with Random Forest Classifier, plotted ROC-AUC curves, and serialized models with Pickle."
            ]
        },
        {
            "company": "FLIPROBO TECHNOLOGY",
            "position": "Data Scientist",
            "location": "Bengaluru, India",
            "start_date": "Mar 2022",
            "end_date": "Feb 2023",
            "description": [
                "1. Insurance Claims - Fraud Detection",
                "Exploratory Data Analysis, pre-processing, removed skewness using log transformation, treated outliers using capping/flooring technique.",
                "Used various ML algorithms XGBOOST, SVM, Logistic regression, Random Forest and Decision trees.",
                "Achieved 89% accuracy using stacking of models."
            ]
        },
        {
            "company": "SATYENDRA AGRO-FOODS PRIVATE LIMITED",
            "position": "Data Scientist",
            "location": "Kolkata, India",
            "start_date": "Mar 2018",
            "end_date": "Oct 2021",
            "description": [
                "Created sales dashboards using Tableau and Power BI to analyze historical data and forecast sales trends for North and East zones.",
                "Applied machine learning classification algorithms to predict agrofood sales in North and East zones of Bihar.",
                "Connected to various data sources, edited data schemas, and performed data sorting, filtering, and grouping."
            ]
        },
        {
            "company": "M N DASTUR CO. PVT. LTD",
            "position": "Data Scientist",
            "location": "Kalinganagar, Odisha",
            "start_date": "Mar 2015",
            "end_date": "Nov 2017",
            "description": [
                "Collected data from disparate data sources using complex SQL/MySQL queries.",
                "Built interactive dashboards based on KPIs and other corporate performance indicators using Power BI.",
                "Developed machine learning models to detect welding faults and predict the quantity of steel required for the 6MTPA construction project."
            ]
        },
        {
            "company": "M N Dastur Pvt. Ltd",
            "position": "Senior Engineer",
            "location": "Kalinganagar, Odisha",
            "start_date": "Jun 2011",
            "end_date": "Mar 2015",
            "description": [
                "Supervised the erection of various complex steel structures and mechanical equipment.",
                "Ensured compliance and quality assurance through regular pre- and post-erection inspections.",
                "Maintained strict engineering protocols, documentation, and quality tolerance levels.",
                "Enforced site safety norms and standards, maintaining a zero-safety-tolerance limit."
            ]
        },
        {
            "company": "Petro Carbon and Chemicals Pvt Ltd",
            "position": "Production Engineer",
            "location": "Haldia, India",
            "start_date": "Nov 2009",
            "end_date": "Jan 2011",
            "description": [
                "Inspected raw materials prior to production to ensure quality and specification conformity.",
                "Operated and maintained rotary kiln, cooler, and stack systems in a continuous production environment."
            ]
        }
    ],
    "education": [
        {
            "institution": "Data Trained Education",
            "degree": "PG in Data Science (IBM collaboration)",
            "location": "Bengaluru, India",
            "start_date": "2021",
            "end_date": "2022",
            "description": "Completed Post Graduate Diploma in Data Science and Machine Learning from Data Trained in collaboration with IBM."
        },
        {
            "institution": "HALDIA INSTITUTE OF TECHNOLOGY",
            "degree": "B. Tech",
            "location": "Haldia, India",
            "start_date": "May 2005",
            "end_date": "Jul 2009",
            "description": "Graduated with B. Tech in Engineering."
        }
    ],
    "skills": [
        {
            "category": "Programming & Databases",
            "skills": ["Python", "SQL", "MySQL"]
        },
        {
            "category": "Machine Learning & AI",
            "skills": ["Machine Learning", "Generative AI", "RAG Architecture", "LLMs (GPT-4, Llama)", "Variance Inflation Factor (VIF)", "SMOTE", "Random Forest Classifier", "Pickle"]
        },
        {
            "category": "Data Analysis & BI",
            "skills": ["Pandas", "Microsoft Power BI", "Tableau", "Matplotlib", "Seaborn", "SAS VIYA 9.4", "SAS VDMML", "SAS Enterprise Guide"]
        },
        {
            "category": "Cloud & Infrastructure",
            "skills": ["AWS", "Azure", "Git"]
        }
    ],
    "projects": [
        {
            "name": "Rain Prediction – Weather Forecasting",
            "role": "Data Scientist",
            "link": "",
            "description": "Utilized 10 years of daily weather observations from different locations in Australia to predict whether it will rain tomorrow and forecast rainfall in mm. Implemented both classification and regression methods. Followed full EDA workflow, applied feature engineering (extracted Year, Month, Day from Date), imputed nulls, checked skewness using Power Transformer, addressed multicollinearity using VIF, balanced dataset using SMOTE, reduced dimensionality using PCA, and hyper-tuned Random Forest Regressor using Grid Search CV (achieved 93% accuracy)."
        },
        {
            "name": "Customer Churn Prediction",
            "role": "Data Scientist",
            "link": "",
            "description": "Designed a machine learning workflow to predict telecom customer churn. Fetched and stored transactional data in a master database, analyzed data patterns and trends using visualization tools, split the dataset into features and labels, trained classification models, and deployed the inference pipeline on AWS Cloud."
        }
    ],
    "certifications": [
        {
            "name": "Post Graduate Diploma in Data Science",
            "issuer": "DataTrained in collaboration with IBM",
            "date": "2022"
        }
    ],
    "languages": [
        {
            "name": "English",
            "proficiency": "Professional Working"
        },
        {
            "name": "Hindi",
            "proficiency": "Native"
        },
        {
            "name": "Bengali",
            "proficiency": "Conversational"
        }
    ],
    "accomplishments": [
        "Trained various Classification models to predict if it will rain tomorrow; Random Forest Classifier achieved 93% accuracy and 84% cross-validation score.",
        "Built a RAG-based chatbot and deployed it on Microsoft Azure for interactive shipping queries and shipment tracking.",
        "Constructed a predictive model that forecasts the probability of default for loan risk assessment at UCO Bank with 85% accuracy."
    ]
}

# Sidebar - API Keys & Template Settings
with st.sidebar:
    st.image("https://resume.io/assets/images/logo/logo-square.png", width=60)
    st.markdown("### Configuration & Settings")
    
    env_tavily = os.environ.get("TAVILY_API_KEY", "")
    env_groq = os.environ.get("GROQ_API_KEY", "")
    env_openai = os.environ.get("OPENAI_API_KEY", "")
    
    # Initialize keys from session state or environment defaults
    tavily_key = st.session_state.get("tavily_key") or env_tavily
    groq_key = st.session_state.get("groq_key") or env_groq
    openai_key = st.session_state.get("openai_key") or env_openai
    
    st.session_state["tavily_key"] = tavily_key
    st.session_state["groq_key"] = groq_key
    st.session_state["openai_key"] = openai_key
    
    # If keys are defined on the backend/env, hide the manual inputs
    if not (env_groq or env_openai):
        st.markdown("#### 🔑 API Credentials")
        tavily_key_input = st.text_input("Tavily API Key", type="password", 
                                   value=tavily_key,
                                   help="Required for web search enrichment.")
        groq_key_input = st.text_input("Groq API Key (Primary LLM)", type="password", 
                                 value=groq_key,
                                 help="Runs Llama-3.3-70b-versatile. Recommended.")
        openai_key_input = st.text_input("OpenAI API Key (Fallback LLM)", type="password", 
                                   value=openai_key,
                                   help="Used as fallback if Groq fails or rate limits.")
        
        # If user changed inputs, update state and local variables
        if tavily_key_input != tavily_key:
            st.session_state["tavily_key"] = tavily_key_input
            tavily_key = tavily_key_input
        if groq_key_input != groq_key:
            st.session_state["groq_key"] = groq_key_input
            groq_key = groq_key_input
        if openai_key_input != openai_key:
            st.session_state["openai_key"] = openai_key_input
            openai_key = openai_key_input
            
        # Status Indicators
        st.markdown("#### ⚡ Connection Status")
        if groq_key:
            st.markdown("🟢 **Groq LLM:** Configured")
        else:
            st.markdown("🔴 **Groq LLM:** Missing")
            
        if openai_key:
            st.markdown("🟢 **OpenAI (Fallback):** Configured")
        else:
            st.markdown("🟡 **OpenAI (Fallback):** Missing (No Fallback)")
            
        if tavily_key:
            st.markdown("🟢 **Tavily Search:** Configured")
        else:
            st.markdown("🔴 **Tavily Search:** Missing")
    else:
        st.markdown("#### ⚡ System Status")
        st.markdown("🟢 **AI Services Active (Backend)**")

    st.markdown("---")
    
    st.markdown("#### 🎨 Design Settings")
    template_option = st.selectbox(
        "Choose Resume Template",
        ("Dublin (Two Column)", "Toronto (Accent Header)", "Stockholm (Minimalist Elegant)")
    )
    
    # Translate option to name
    template_map = {
        "Dublin (Two Column)": "Dublin",
        "Toronto (Accent Header)": "Toronto",
        "Stockholm (Minimalist Elegant)": "Stockholm"
    }
    active_template = template_map[template_option]

# Title Block Card
st.markdown("""
<div class="header-card">
    <h1>📄 Resume.io AI Builder & Co-Pilot</h1>
    <p>Build and optimize professional-grade resumes to resume.io standards using deep agents powered by LangGraph & LangChain.</p>
</div>
""", unsafe_allow_html=True)

# Action button to load sample data
col_load, col_clear = st.columns([1, 4])
with col_load:
    if st.button("📥 Load Sample Data", help="Load a pre-built vague developer resume to test the AI refiner."):
        st.session_state.resume_data = sample_data
        st.success("Sample data loaded! Go to the 'AI Co-Pilot' tab to refine it.")
with col_clear:
    if st.button("🧹 Reset Form", help="Clear all fields to write your own resume."):
        st.session_state.resume_data = {
            "personal_info": {
                "full_name": "", 
                "professional_title": "", 
                "email": "", 
                "phone": "", 
                "location": "", 
                "linkedin": "", 
                "github": "", 
                "website": "",
                "birth_info": "",
                "nationality": "",
                "hobbies": "",
                "photo_base64": None,
                "power_statement": ""
            },
            "summary": "",
            "experience": [],
            "education": [],
            "skills": [],
            "projects": [],
            "certifications": [],
            "languages": [],
            "accomplishments": []
        }
        st.rerun()

def apply_ai_suggestion(suggestion):
    section = suggestion.get("section")
    index = suggestion.get("index")
    key = suggestion.get("key")
    bullet_index = suggestion.get("bullet_index")
    suggested = suggestion.get("suggested")
    
    resume_data = st.session_state.resume_data
    
    if section == "summary":
        resume_data["summary"] = suggested
    elif section == "personal_info" and key:
        if "personal_info" not in resume_data:
            resume_data["personal_info"] = {}
        resume_data["personal_info"][key] = suggested
    elif section == "experience" and index is not None:
        if index < len(resume_data.get("experience", [])):
            if bullet_index is not None:
                desc = resume_data["experience"][index].get("description", [])
                if bullet_index < len(desc):
                    desc[bullet_index] = suggested
                else:
                    desc.append(suggested)
                resume_data["experience"][index]["description"] = desc
            elif key:
                resume_data["experience"][index][key] = suggested
            else:
                if isinstance(suggested, list):
                    resume_data["experience"][index]["description"] = suggested
                elif isinstance(suggested, str):
                    resume_data["experience"][index]["description"] = [s.strip() for s in suggested.split("\n") if s.strip()]
    elif section == "skills" and index is not None:
        if index < len(resume_data.get("skills", [])):
            if key == "category":
                resume_data["skills"][index]["category"] = suggested
            elif key == "skills":
                if isinstance(suggested, list):
                    resume_data["skills"][index]["skills"] = suggested
                elif isinstance(suggested, str):
                    resume_data["skills"][index]["skills"] = [s.strip() for s in suggested.split(",") if s.strip()]
    elif section == "accomplishments":
        if index is not None:
            if "accomplishments" not in resume_data:
                resume_data["accomplishments"] = []
            if index < len(resume_data["accomplishments"]):
                resume_data["accomplishments"][index] = suggested
            else:
                resume_data["accomplishments"].append(suggested)
        else:
            if isinstance(suggested, list):
                resume_data["accomplishments"] = suggested
            elif isinstance(suggested, str):
                resume_data["accomplishments"] = [s.strip() for s in suggested.split("\n") if s.strip()]
    elif section == "projects" and index is not None:
        if index < len(resume_data.get("projects", [])):
            if key:
                resume_data["projects"][index][key] = suggested
    elif section == "certifications" and index is not None:
        if index < len(resume_data.get("certifications", [])):
            if key:
                resume_data["certifications"][index][key] = suggested
    elif section == "languages" and index is not None:
        if index < len(resume_data.get("languages", [])):
            if key:
                resume_data["languages"][index][key] = suggested

    st.session_state.resume_data = resume_data

# Define main application tabs
tab_fill, tab_copilot, tab_preview = st.tabs([
    "✏️ Fill & Edit Details",
    "🤖 AI Assistant Co-Pilot",
    "👁️ Preview & Download"
])

# ================= TAB 1: FILL DETAILS =================
with tab_fill:
    st.subheader("Edit Your Resume Information")
    
    # Personal Info Section
    with st.expander("👤 Contact & Personal Details", expanded=True):
        col1, col2 = st.columns(2)
        pi = st.session_state.resume_data["personal_info"]
        
        pi["full_name"] = col1.text_input("Full Name", value=pi.get("full_name", ""))
        pi["professional_title"] = col2.text_input("Professional Title", value=pi.get("professional_title", ""), placeholder="e.g. Lead Software Architect")
        
        pi["email"] = col1.text_input("Email", value=pi.get("email", ""))
        pi["phone"] = col2.text_input("Phone Number", value=pi.get("phone", ""))
        
        pi["location"] = col1.text_input("Location (City, State/Country)", value=pi.get("location", ""), placeholder="e.g. Bengaluru, India")
        pi["linkedin"] = col2.text_input("LinkedIn Profile URL", value=pi.get("linkedin", ""))
        
        pi["github"] = col1.text_input("GitHub Profile URL", value=pi.get("github", ""))
        pi["website"] = col2.text_input("Portfolio Website URL", value=pi.get("website", ""))
        
        pi["birth_info"] = col1.text_input("Date / Place of Birth", value=pi.get("birth_info", ""), placeholder="e.g. 05/03/1984, Bokaro")
        pi["nationality"] = col2.text_input("Nationality", value=pi.get("nationality", ""), placeholder="e.g. Indian")
        
        pi["hobbies"] = col1.text_input("Hobbies", value=pi.get("hobbies", ""), placeholder="e.g. Website Development, Reading")
        pi["power_statement"] = col2.text_input("Power Statement / Tagline", value=pi.get("power_statement", ""), placeholder="e.g. Leveraging AI to solve complex challenges")
        
        st.markdown("##### 📸 Profile Photo")
        uploaded_file = st.file_uploader("Upload Profile Photo (JPG/PNG)", type=["jpg", "jpeg", "png"], help="Upload a photo to display on your resume.")
        if uploaded_file is not None:
            bytes_data = uploaded_file.read()
            base64_image = base64.b64encode(bytes_data).decode("utf-8")
            mime_type = uploaded_file.type
            pi["photo_base64"] = f"data:{mime_type};base64,{base64_image}"
            st.success("Photo uploaded successfully!")
            
        if pi.get("photo_base64"):
            col_img, col_btn = st.columns([1, 4])
            with col_img:
                st.image(pi["photo_base64"], width=80, caption="Preview")
            with col_btn:
                st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
                if st.button("🗑️ Remove Photo", key="remove_photo_btn"):
                    pi["photo_base64"] = None
                    st.session_state.resume_data["personal_info"] = pi
                    st.rerun()
        
        st.session_state.resume_data["personal_info"] = pi

    # Professional Summary Section
    with st.expander("📝 Professional Summary"):
        st.session_state.resume_data["summary"] = st.text_area(
            "Write a brief summary of your skills and career highlights", 
            value=st.session_state.resume_data.get("summary", ""),
            height=120,
            placeholder="A short 2-4 sentence pitch describing your strengths..."
        )

    # Work Experience Section
    with st.expander("💼 Work Experience"):
        exp_list = st.session_state.resume_data.get("experience", [])
        
        # Display existing entries and let user edit
        updated_exp = []
        for idx, exp in enumerate(exp_list):
            st.markdown(f"##### Position {idx+1}")
            col1, col2 = st.columns(2)
            comp = col1.text_input(f"Company Name #{idx+1}", value=exp.get("company", ""), key=f"comp_{idx}")
            pos = col2.text_input(f"Job Title #{idx+1}", value=exp.get("position", ""), key=f"pos_{idx}")
            
            col3, col4, col5 = st.columns(3)
            loc = col3.text_input(f"Location #{idx+1}", value=exp.get("location", ""), key=f"loc_{idx}")
            s_date = col4.text_input(f"Start Date #{idx+1}", value=exp.get("start_date", ""), placeholder="e.g. Jan 2021", key=f"s_date_{idx}")
            e_date = col5.text_input(f"End Date #{idx+1}", value=exp.get("end_date", ""), placeholder="e.g. Present", key=f"e_date_{idx}")
            
            # Bullet points
            desc_text = st.text_area(
                f"Bullet Points (one achievement per line) #{idx+1}", 
                value="\n".join(exp.get("description", [])),
                height=100,
                key=f"desc_{idx}",
                help="Enter each bullet point on a new line."
            )
            desc_bullets = [line.strip() for line in desc_text.split("\n") if line.strip()]
            
            if st.button(f"🗑️ Remove Position #{idx+1}", key=f"rem_exp_{idx}"):
                # Remove this item and rerun
                exp_list.pop(idx)
                st.session_state.resume_data["experience"] = exp_list
                st.rerun()
                
            updated_exp.append({
                "company": comp,
                "position": pos,
                "location": loc,
                "start_date": s_date,
                "end_date": e_date,
                "description": desc_bullets
            })
            st.markdown("---")
            
        # Add new entry
        if st.button("➕ Add Work Experience"):
            exp_list.append({"company": "", "position": "", "location": "", "start_date": "", "end_date": "", "description": []})
            st.session_state.resume_data["experience"] = exp_list
            st.rerun()
            
        st.session_state.resume_data["experience"] = updated_exp

    # Accomplishments Section
    with st.expander("🏆 Accomplishments"):
        acc_list = st.session_state.resume_data.get("accomplishments", [])
        
        # Display existing entries and let user edit
        updated_acc = []
        for idx, acc in enumerate(acc_list):
            col_acc, col_btn = st.columns([5, 1])
            val = col_acc.text_input(f"Accomplishment #{idx+1}", value=acc, key=f"acc_{idx}")
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            if col_btn.button(f"🗑️ Remove #{idx+1}", key=f"rem_acc_{idx}"):
                acc_list.pop(idx)
                st.session_state.resume_data["accomplishments"] = acc_list
                st.rerun()
            updated_acc.append(val)
            
        # Add new entry
        if st.button("➕ Add Accomplishment"):
            acc_list.append("")
            st.session_state.resume_data["accomplishments"] = acc_list
            st.rerun()
            
        st.session_state.resume_data["accomplishments"] = updated_acc

    # Education Section
    with st.expander("🎓 Education"):
        edu_list = st.session_state.resume_data.get("education", [])
        
        updated_edu = []
        for idx, edu in enumerate(edu_list):
            st.markdown(f"##### Education {idx+1}")
            col1, col2 = st.columns(2)
            inst = col1.text_input(f"School/University #{idx+1}", value=edu.get("institution", ""), key=f"inst_{idx}")
            deg = col2.text_input(f"Degree/Major #{idx+1}", value=edu.get("degree", ""), key=f"deg_{idx}")
            
            col3, col4, col5 = st.columns(3)
            loc = col3.text_input(f"Location (School) #{idx+1}", value=edu.get("location", ""), key=f"eduloc_{idx}")
            s_date = col4.text_input(f"Start Date (School) #{idx+1}", value=edu.get("start_date", ""), key=f"edus_date_{idx}")
            e_date = col5.text_input(f"End Date (School) #{idx+1}", value=edu.get("end_date", ""), key=f"edue_date_{idx}")
            
            details = st.text_area(f"Additional Details #{idx+1}", value=edu.get("description", ""), height=60, placeholder="e.g. GPA, Honors, Key Coursework", key=f"edudet_{idx}")
            
            if st.button(f"🗑️ Remove Education #{idx+1}", key=f"rem_edu_{idx}"):
                edu_list.pop(idx)
                st.session_state.resume_data["education"] = edu_list
                st.rerun()
                
            updated_edu.append({
                "institution": inst,
                "degree": deg,
                "location": loc,
                "start_date": s_date,
                "end_date": e_date,
                "description": details
            })
            st.markdown("---")
            
        if st.button("➕ Add Education"):
            edu_list.append({"institution": "", "degree": "", "location": "", "start_date": "", "end_date": "", "description": ""})
            st.session_state.resume_data["education"] = edu_list
            st.rerun()
            
        st.session_state.resume_data["education"] = updated_edu

    # Skills Section
    with st.expander("🛠️ Skills (Categorized)"):
        skills_list = st.session_state.resume_data.get("skills", [])
        
        updated_skills = []
        for idx, group in enumerate(skills_list):
            st.markdown(f"##### Skill Group {idx+1}")
            col1, col2 = st.columns([1, 2])
            cat = col1.text_input(f"Category Name #{idx+1}", value=group.get("category", ""), placeholder="e.g. Programming Languages", key=f"cat_{idx}")
            s_list_str = col2.text_input(f"Skills (comma separated) #{idx+1}", value=", ".join(group.get("skills", [])), placeholder="e.g. Python, SQL, C++", key=f"slist_{idx}")
            
            skills_arr = [s.strip() for s in s_list_str.split(",") if s.strip()]
            
            if st.button(f"🗑️ Remove Skill Group #{idx+1}", key=f"rem_sk_{idx}"):
                skills_list.pop(idx)
                st.session_state.resume_data["skills"] = skills_list
                st.rerun()
                
            updated_skills.append({
                "category": cat,
                "skills": skills_arr
            })
            st.markdown("---")
            
        if st.button("➕ Add Skill Group"):
            skills_list.append({"category": "", "skills": []})
            st.session_state.resume_data["skills"] = skills_list
            st.rerun()
            
        st.session_state.resume_data["skills"] = updated_skills

    # Projects Section
    with st.expander("🚀 Side Projects"):
        proj_list = st.session_state.resume_data.get("projects", [])
        
        updated_proj = []
        for idx, proj in enumerate(proj_list):
            st.markdown(f"##### Project {idx+1}")
            col1, col2, col3 = st.columns([2, 1, 2])
            p_name = col1.text_input(f"Project Name #{idx+1}", value=proj.get("name", ""), key=f"pname_{idx}")
            p_role = col2.text_input(f"Role #{idx+1}", value=proj.get("role", ""), placeholder="e.g. Sole Developer", key=f"prole_{idx}")
            p_link = col3.text_input(f"Project Link #{idx+1}", value=proj.get("link", ""), placeholder="e.g. GitHub URL", key=f"plink_{idx}")
            
            p_desc = st.text_area(f"Project Description #{idx+1}", value=proj.get("description", ""), height=70, key=f"pdesc_{idx}")
            
            if st.button(f"🗑️ Remove Project #{idx+1}", key=f"rem_pr_{idx}"):
                proj_list.pop(idx)
                st.session_state.resume_data["projects"] = proj_list
                st.rerun()
                
            updated_proj.append({
                "name": p_name,
                "role": p_role,
                "link": p_link,
                "description": p_desc
            })
            st.markdown("---")
            
        if st.button("➕ Add Project"):
            proj_list.append({"name": "", "role": "", "link": "", "description": ""})
            st.session_state.resume_data["projects"] = proj_list
            st.rerun()
            
        st.session_state.resume_data["projects"] = updated_proj

    # Certifications Section
    with st.expander("📜 Certifications"):
        cert_list = st.session_state.resume_data.get("certifications", [])
        
        updated_certs = []
        for idx, cert in enumerate(cert_list):
            st.markdown(f"##### Certification {idx+1}")
            col1, col2, col3 = st.columns(3)
            c_name = col1.text_input(f"Certification Name #{idx+1}", value=cert.get("name", ""), key=f"cname_{idx}")
            c_iss = col2.text_input(f"Issuing Organization #{idx+1}", value=cert.get("issuer", ""), key=f"ciss_{idx}")
            c_date = col3.text_input(f"Date Earned #{idx+1}", value=cert.get("date", ""), key=f"cdate_{idx}")
            
            if st.button(f"🗑️ Remove Certification #{idx+1}", key=f"rem_ce_{idx}"):
                cert_list.pop(idx)
                st.session_state.resume_data["certifications"] = cert_list
                st.rerun()
                
            updated_certs.append({
                "name": c_name,
                "issuer": c_iss,
                "date": c_date
            })
            st.markdown("---")
            
        if st.button("➕ Add Certification"):
            cert_list.append({"name": "", "issuer": "", "date": ""})
            st.session_state.resume_data["certifications"] = cert_list
            st.rerun()
            
        st.session_state.resume_data["certifications"] = updated_certs

    # Languages Section
    with st.expander("🗣️ Languages"):
        lang_list = st.session_state.resume_data.get("languages", [])
        
        updated_langs = []
        for idx, lang in enumerate(lang_list):
            st.markdown(f"##### Language {idx+1}")
            col1, col2 = st.columns(2)
            l_name = col1.text_input(f"Language Name #{idx+1}", value=lang.get("name", ""), key=f"lname_{idx}")
            l_prof = col2.text_input(f"Proficiency Level #{idx+1}", value=lang.get("proficiency", ""), placeholder="e.g. Native, Fluent, Conversational", key=f"lprof_{idx}")
            
            if st.button(f"🗑️ Remove Language #{idx+1}", key=f"rem_la_{idx}"):
                lang_list.pop(idx)
                st.session_state.resume_data["languages"] = lang_list
                st.rerun()
                
            updated_langs.append({
                "name": l_name,
                "proficiency": l_prof
            })
            st.markdown("---")
            
        if st.button("➕ Add Language"):
            lang_list.append({"name": "", "proficiency": ""})
            st.session_state.resume_data["languages"] = lang_list
            st.rerun()
            
        st.session_state.resume_data["languages"] = updated_langs

# ================= TAB 2: AI CO-PILOT ASSISTANT =================
with tab_copilot:
    st.subheader("🤖 AI Assistant Co-Pilot & Enhancer")
    
    subtab_enhance, subtab_rewrite = st.tabs([
        "✨ Auto-Enhancement Suggestions",
        "🚀 Full Deep Agent Rewrite"
    ])
    
    with subtab_enhance:
        st.write("### ✨ Auto-Enhancement Suggestions")
        st.write("Generate specific improvement suggestions for your current resume content, review them side-by-side, and apply them with one click.")
        
        # Session state to store suggestions list
        if "ai_suggestions" not in st.session_state:
            st.session_state.ai_suggestions = []
            
        col_gen, col_clear_sug = st.columns([1, 4])
        with col_gen:
            if st.button("🔍 Generate AI Suggestions", key="gen_sug_btn", disabled=not (groq_key or openai_key)):
                with st.spinner("🤖 Analyzing resume data and generating suggestions..."):
                    try:
                        suggestions = generate_resume_suggestions(
                            st.session_state.resume_data,
                            groq_api_key=groq_key,
                            openai_api_key=openai_key
                        )
                        st.session_state.ai_suggestions = suggestions
                        if suggestions:
                            st.success(f"Generated {len(suggestions)} suggestions!")
                        else:
                            st.info("No improvements found! Your resume looks great.")
                    except Exception as e:
                        st.error(f"Error generating suggestions: {e}")
                        
        with col_clear_sug:
            if st.button("🧹 Clear Suggestions", key="clear_sug_btn"):
                st.session_state.ai_suggestions = []
                st.rerun()
                
        if st.session_state.ai_suggestions:
            st.markdown("---")
            st.write("#### Available Suggestions")
            
            for idx, sug in enumerate(st.session_state.ai_suggestions):
                with st.container():
                    st.markdown(f"##### 💡 {sug.get('title', 'Resume Suggestion')}")
                    st.markdown(f"**Section:** `{sug.get('section', '')}` | **Reason:** {sug.get('reason', '')}")
                    
                    col_before, col_after = st.columns(2)
                    with col_before:
                        st.caption("Before (Current)")
                        original_text = sug.get('original', '')
                        if isinstance(original_text, list):
                            original_text = "\n".join([f"- {item}" for item in original_text])
                        st.text_area("Original Value", value=original_text, height=100, disabled=True, key=f"orig_val_{idx}")
                        
                    with col_after:
                        st.caption("After (Suggested)")
                        suggested_text = sug.get('suggested', '')
                        if isinstance(suggested_text, list):
                            suggested_text = "\n".join([f"- {item}" for item in suggested_text])
                        st.text_area("Suggested Value", value=suggested_text, height=100, disabled=True, key=f"sug_val_{idx}")
                        
                    if st.button("⚡ Apply Suggestion", key=f"apply_sug_{idx}"):
                        apply_ai_suggestion(sug)
                        st.session_state.ai_suggestions.pop(idx)
                        st.success(f"Applied suggestion: '{sug.get('title')}'!")
                        st.rerun()
                        
                    st.markdown("---")
        else:
            st.info("Click 'Generate AI Suggestions' to analyze your resume and get recommendations.")
            
    with subtab_rewrite:
        st.write("### 🚀 Full Deep Agent Rewrite")
        st.markdown("""
        Use the deep agent to automatically rewrite, polish, and structure your details into a premium, resume.io-grade resume.
        The agent will evaluate your target job requirements, rewrite bullet points with powerful action verbs, optimize for ATS, and add metrics.
        """)
        
        st.markdown("### What are you optimizing for?")
        copilot_prompt = st.text_area(
            "Enter optimization instructions (target role, specific guidelines, etc.)",
            value="Optimize this resume for a Senior AI & Full-Stack Engineer role at a top-tier tech company. Search for standard skills and keywords for this role, and rewrite the experience bullet points using the XYZ formula (Quantifiable metrics, action verbs). Make the summary sound highly professional.",
            height=100,
            key="rewrite_instructions_input"
        )
        
        # Warning message if keys are missing
        if not (groq_key or openai_key):
            st.warning("⚠️ Please configure a Groq API Key or OpenAI API Key in the sidebar to activate the AI Co-Pilot.")
        if not tavily_key:
            st.info("ℹ️ Configuring a Tavily API Key is highly recommended so the AI can search for standard requirements for your target role.")
            
        if st.button("🚀 Run Deep Agent Refinement", disabled=not (groq_key or openai_key), key="run_deep_refine_btn"):
            with st.spinner("🤖 Deep Agent starting execution... Running planning, researching target role keywords via Tavily, and rewriting sections... Please wait."):
                try:
                    # Perform the agent rewrite
                    refined_resume = refine_resume_with_agent(
                        current_resume_dict=st.session_state.resume_data,
                        refinement_instructions=copilot_prompt,
                        groq_api_key=groq_key,
                        openai_api_key=openai_key,
                        tavily_api_key=tavily_key
                    )
                    
                    # Update session state with refined resume
                    st.session_state.resume_data = refined_resume.model_dump()
                    st.success("🎉 Resume successfully refined by the AI Agent! Check the 'Fill & Edit Details' or 'Preview & Download' tab.")
                    st.balloons()
                except Exception as e:
                    st.error(f"Error during agent execution: {str(e)}")
                    st.info("Ensure your API keys are correct. If you hit Groq rate limits, try configuring the OpenAI fallback key.")

# ================= TAB 3: PREVIEW & DOWNLOAD =================
with tab_preview:
    st.subheader("👁️ Live Resume Preview")
    
    # Try parsing the state dictionary back into a ResumeData object for rendering
    try:
        current_resume_data = ResumeData(**st.session_state.resume_data)
        
        # HTML Rendering
        html_code = render_html_resume(current_resume_data, active_template)
        
        # Display Preview in Iframe
        st.markdown(f"**Template: {template_option}**")
        st.components.v1.html(html_code, height=650, scrolling=True)
        
        st.markdown("### 📥 Export & Download Options")
        
        col_html, col_pdf, col_excel = st.columns(3)
        
        with col_html:
            st.markdown("#### 🌐 HTML Format")
            st.markdown("Download as a standalone, responsive HTML file that displays correctly on any browser.")
            html_link = get_html_download_link(html_code, filename=f"{current_resume_data.personal_info.full_name.replace(' ', '_')}_resume.html", link_text="📥 Download HTML File")
            st.markdown(html_link, unsafe_allow_html=True)
            
        with col_pdf:
            st.markdown("#### 📄 PDF Format")
            st.markdown("Save as a high-quality PDF. Toggles a print layout containing proper paper breaks and fonts.")
            print_link = get_html_print_link(html_code, link_text="🖨️ Open Print & Save as PDF")
            st.markdown(print_link, unsafe_allow_html=True)
            st.markdown("<small>*Tip: In the print dialog, set layout to Portrait, margins to Default/None, and turn on 'Background Graphics'.*</small>", unsafe_allow_html=True)
            
        with col_excel:
            st.markdown("#### 📊 Excel Format")
            st.markdown("Export all resume details in a beautiful, structured multi-sheet spreadsheet for custom editing.")
            try:
                excel_bytes = export_to_excel(current_resume_data)
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_bytes,
                    file_name=f"{current_resume_data.personal_info.full_name.replace(' ', '_')}_resume.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error generating Excel file: {e}")
                
    except Exception as e:
        st.error(f"Could not parse resume data to generate preview: {e}")
        st.info("Make sure your resume details are not empty. Go to the 'Fill & Edit Details' tab and add content or click 'Load Sample Data'.")
