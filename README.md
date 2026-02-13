# Cencora Launch Readiness Assessment System

A self-assessment tool for leadership development programmes, measuring readiness across four key indicators before and after programme delivery.

## Features

- **Pre/Post Assessment**: Participants complete self-assessments before and after the programme
- **32-Item Framework**: Covering Self-Readiness, Practical Readiness, Professional Readiness, and Team Readiness
- **Three Report Types**:
  - **Baseline Report**: Individual pre-programme snapshot
  - **Progress Report**: Individual pre vs post comparison with cohort benchmarks
  - **Impact Report**: Cohort-level summary with AI-powered theme extraction
- **Cloud Database**: Persistent storage with Turso (SQLite edge database)
- **AI Theme Analysis**: Claude API integration for qualitative response analysis

## Technology Stack

- **Frontend**: Streamlit
- **Database**: Turso (libSQL)
- **Reports**: python-docx, matplotlib
- **AI**: Anthropic Claude API

## Setup

1. Clone this repository
2. Install dependencies: `pip install -r requirements.txt`
3. Configure Streamlit secrets (see below)
4. Run: `streamlit run app.py`

## Configuration

Add to `.streamlit/secrets.toml` or Streamlit Cloud secrets:

```toml
[turso]
url = "libsql://your-database.turso.io"
token = "your-turso-token"

[anthropic]
api_key = "sk-ant-your-api-key"

[app]
base_url = "https://your-app.streamlit.app"
```

## Framework

The assessment measures 32 statements across 4 indicators:

| Indicator | Items | Focus |
|-----------|-------|-------|
| Self-Readiness | 1-6 | Personal awareness, values, presence |
| Practical Readiness | 7-14 | Time, delegation, feedback |
| Professional Readiness | 15-22 | Communication, trust, accountability |
| Team Readiness | 23-30 | HR, safety, change, resilience |
| Overall | 31-32 | Overall readiness confidence |

Each item is tagged with a focus area: Knowledge, Awareness, Confidence, or Behaviour (BACK).

## Licence

Proprietary - The Development Catalyst
