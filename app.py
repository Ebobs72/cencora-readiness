#!/usr/bin/env python3
"""
Cencora Launch Readiness Assessment System

Main Streamlit application with admin interface for managing cohorts,
participants, and generating reports.
"""

import streamlit as st
from datetime import datetime

from database import Database
from framework import (
    INDICATORS, INDICATOR_DESCRIPTIONS, INDICATOR_COLOURS,
    ITEMS, OPEN_QUESTIONS_PRE, OPEN_QUESTIONS_POST, RATING_SCALE
)

# Page configuration
st.set_page_config(
    page_title="Launch Readiness",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Cencora branding
st.markdown("""
<style>
    .main-header {
        color: #461E96;
        font-size: 2rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        color: #E6008C;
        font-size: 1.2rem;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F5F5F5;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #461E96;
    }
    .status-complete {
        color: #00DC8C;
        font-weight: bold;
    }
    .status-pending {
        color: #FFA400;
        font-weight: bold;
    }
    .stButton button {
        background-color: #461E96;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# Initialize database
@st.cache_resource
def get_database():
    return Database()

db = get_database()


def main():
    """Main application entry point."""
    
    # Check if this is an assessment link
    query_params = st.query_params
    if 'token' in query_params:
        show_assessment_form(query_params['token'])
        return
    
    # Otherwise show admin interface
    show_admin_interface()


def show_admin_interface():
    """Display the admin interface."""
    
    st.markdown('<p class="main-header">🚀 Launch Readiness</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Assessment Administration</p>', unsafe_allow_html=True)
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Go to",
        ["Overview", "Manage Cohorts", "Manage Participants", "Generate Reports", "Settings"]
    )
    
    if page == "Overview":
        show_overview()
    elif page == "Manage Cohorts":
        show_cohort_management()
    elif page == "Manage Participants":
        show_participant_management()
    elif page == "Generate Reports":
        show_report_generation()
    elif page == "Settings":
        show_settings()


def show_overview():
    """Display overview dashboard."""
    
    st.header("Dashboard")
    
    cohorts = db.get_all_cohorts()
    
    if not cohorts:
        st.info("No cohorts created yet. Go to 'Manage Cohorts' to create your first cohort.")
        return
    
    # Summary metrics
    total_participants = 0
    total_pre_complete = 0
    total_post_complete = 0
    
    for cohort in cohorts:
        participants = db.get_participants_for_cohort(cohort['id'])
        total_participants += len(participants)
        total_pre_complete += sum(1 for p in participants if p.get('pre_completed'))
        total_post_complete += sum(1 for p in participants if p.get('post_completed'))
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Cohorts", len(cohorts))
    with col2:
        st.metric("Total Participants", total_participants)
    with col3:
        st.metric("PRE Assessments Complete", total_pre_complete)
    with col4:
        st.metric("POST Assessments Complete", total_post_complete)
    
    st.divider()
    
    # Cohort summary table
    st.subheader("Cohorts at a Glance")
    
    for cohort in cohorts:
        participants = db.get_participants_for_cohort(cohort['id'])
        pre_done = sum(1 for p in participants if p.get('pre_completed'))
        post_done = sum(1 for p in participants if p.get('post_completed'))
        
        with st.expander(f"**{cohort['name']}** - {len(participants)} participants"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write(f"**Programme:** {cohort.get('programme', 'Launch Readiness')}")
            with col2:
                st.write(f"**PRE Complete:** {pre_done}/{len(participants)}")
            with col3:
                st.write(f"**POST Complete:** {post_done}/{len(participants)}")
            
            if cohort.get('description'):
                st.write(f"*{cohort['description']}*")


def show_cohort_management():
    """Manage cohorts."""
    
    st.header("Manage Cohorts")
    
    # Create new cohort
    st.subheader("Create New Cohort")
    
    with st.form("new_cohort"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Cohort Name*", placeholder="e.g., Cohort 1 - January 2026")
            programme = st.text_input("Programme Name", value="Launch Readiness")
        with col2:
            start_date = st.date_input("Start Date")
            end_date = st.date_input("End Date")
        
        description = st.text_area("Description (optional)", placeholder="Any notes about this cohort...")
        
        if st.form_submit_button("Create Cohort", type="primary"):
            if name:
                cohort_id = db.create_cohort(
                    name=name,
                    programme=programme,
                    description=description,
                    start_date=start_date.isoformat() if start_date else None,
                    end_date=end_date.isoformat() if end_date else None
                )
                st.success(f"Cohort '{name}' created successfully!")
                st.rerun()
            else:
                st.error("Please enter a cohort name.")
    
    st.divider()
    
    # List existing cohorts
    st.subheader("Existing Cohorts")
    
    cohorts = db.get_all_cohorts()
    
    if not cohorts:
        st.info("No cohorts created yet.")
        return
    
    for cohort in cohorts:
        participants = db.get_participants_for_cohort(cohort['id'])
        
        with st.expander(f"📁 {cohort['name']} ({len(participants)} participants)"):
            col1, col2, col3 = st.columns([2, 2, 1])
            
            with col1:
                st.write(f"**Programme:** {cohort.get('programme', 'Launch Readiness')}")
                if cohort.get('start_date'):
                    st.write(f"**Dates:** {cohort.get('start_date')} to {cohort.get('end_date', 'TBC')}")
            
            with col2:
                if cohort.get('description'):
                    st.write(f"**Notes:** {cohort['description']}")
            
            with col3:
                if st.button("🗑️ Delete", key=f"del_cohort_{cohort['id']}"):
                    if len(participants) > 0:
                        st.warning(f"This will delete {len(participants)} participants and all their data!")
                    db.delete_cohort(cohort['id'])
                    st.success("Cohort deleted.")
                    st.rerun()


def show_participant_management():
    """Manage participants within cohorts."""
    
    st.header("Manage Participants")
    
    cohorts = db.get_all_cohorts()
    
    if not cohorts:
        st.warning("Create a cohort first before adding participants.")
        return
    
    # Select cohort
    cohort_options = {c['name']: c['id'] for c in cohorts}
    selected_cohort_name = st.selectbox("Select Cohort", options=list(cohort_options.keys()))
    selected_cohort_id = cohort_options[selected_cohort_name]
    
    st.divider()
    
    # Add new participant
    st.subheader("Add Participant")
    
    with st.form("new_participant"):
        col1, col2, col3 = st.columns(3)
        with col1:
            name = st.text_input("Name*", placeholder="e.g., Sarah Mitchell")
        with col2:
            email = st.text_input("Email", placeholder="sarah.mitchell@company.com")
        with col3:
            role = st.text_input("Role", placeholder="e.g., Shift Supervisor")
        
        if st.form_submit_button("Add Participant", type="primary"):
            if name:
                participant_id = db.create_participant(
                    cohort_id=selected_cohort_id,
                    name=name,
                    email=email,
                    role=role
                )
                st.success(f"Participant '{name}' added successfully!")
                st.rerun()
            else:
                st.error("Please enter a name.")
    
    st.divider()
    
    # List participants
    st.subheader("Participants")
    
    participants = db.get_participants_for_cohort(selected_cohort_id)
    
    if not participants:
        st.info("No participants in this cohort yet.")
        return
    
    # Get base URL for assessment links
    try:
        # Try to get the actual app URL
        base_url = st.secrets.get("app", {}).get("base_url", "")
    except:
        base_url = ""
    
    if not base_url:
        base_url = "https://your-app-url.streamlit.app"
    
    for p in participants:
        pre_status = "✅" if p.get('pre_completed') else "⏳"
        post_status = "✅" if p.get('post_completed') else "⏳"
        
        with st.expander(f"{pre_status} {post_status} **{p['name']}** - {p.get('role', 'No role specified')}"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**PRE Assessment**")
                if p.get('pre_completed'):
                    st.write(f"✅ Completed: {p['pre_completed'][:10]}")
                else:
                    pre_link = f"{base_url}?token={p['pre_token']}"
                    st.code(pre_link, language=None)
                    st.caption("Share this link with the participant")
            
            with col2:
                st.write("**POST Assessment**")
                if p.get('post_completed'):
                    st.write(f"✅ Completed: {p['post_completed'][:10]}")
                elif p.get('pre_completed'):
                    post_link = f"{base_url}?token={p['post_token']}"
                    st.code(post_link, language=None)
                    st.caption("Share after programme completion")
                else:
                    st.write("⏳ Complete PRE assessment first")
            
            st.divider()
            
            col1, col2, col3 = st.columns([1, 1, 2])
            with col1:
                if st.button("🗑️ Delete", key=f"del_p_{p['id']}"):
                    db.delete_participant(p['id'])
                    st.success(f"Participant '{p['name']}' deleted.")
                    st.rerun()


def show_report_generation():
    """Generate reports."""
    
    st.header("Generate Reports")
    
    # Import report generator here to avoid circular imports
    from report_generator import ReportGenerator
    
    report_gen = ReportGenerator(db)
    
    cohorts = db.get_all_cohorts()
    
    if not cohorts:
        st.warning("Create a cohort and add participants first.")
        return
    
    # Select cohort
    cohort_options = {c['name']: c['id'] for c in cohorts}
    selected_cohort_name = st.selectbox("Select Cohort", options=list(cohort_options.keys()))
    selected_cohort_id = cohort_options[selected_cohort_name]
    
    participants = db.get_participants_for_cohort(selected_cohort_id)
    
    st.divider()
    
    # Report type selection
    report_type = st.radio(
        "Report Type",
        ["Individual Baseline (PRE)", "Individual Progress (PRE vs POST)", "Cohort Impact Summary"],
        horizontal=True
    )
    
    st.divider()
    
    if report_type == "Individual Baseline (PRE)":
        st.subheader("Generate Baseline Report")
        
        # Filter to those with PRE complete
        eligible = [p for p in participants if p.get('pre_completed')]
        
        if not eligible:
            st.warning("No participants have completed their PRE assessment yet.")
            return
        
        participant_options = {p['name']: p['id'] for p in eligible}
        selected_name = st.selectbox("Select Participant", options=list(participant_options.keys()))
        selected_id = participant_options[selected_name]
        
        if st.button("Generate Baseline Report", type="primary"):
            with st.spinner("Generating report..."):
                try:
                    doc_buffer = report_gen.generate_baseline_report(selected_id)
                    st.download_button(
                        label="📥 Download Baseline Report",
                        data=doc_buffer,
                        file_name=f"Readiness_Baseline_{selected_name.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.success("Report generated successfully!")
                except Exception as e:
                    st.error(f"Error generating report: {e}")
    
    elif report_type == "Individual Progress (PRE vs POST)":
        st.subheader("Generate Progress Report")
        
        # Filter to those with both PRE and POST complete
        eligible = [p for p in participants if p.get('pre_completed') and p.get('post_completed')]
        
        if not eligible:
            st.warning("No participants have completed both PRE and POST assessments yet.")
            return
        
        participant_options = {p['name']: p['id'] for p in eligible}
        selected_name = st.selectbox("Select Participant", options=list(participant_options.keys()))
        selected_id = participant_options[selected_name]
        
        if st.button("Generate Progress Report", type="primary"):
            with st.spinner("Generating report..."):
                try:
                    doc_buffer = report_gen.generate_progress_report(selected_id, selected_cohort_id)
                    st.download_button(
                        label="📥 Download Progress Report",
                        data=doc_buffer,
                        file_name=f"Readiness_Progress_{selected_name.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.success("Report generated successfully!")
                except Exception as e:
                    st.error(f"Error generating report: {e}")
    
    else:  # Cohort Impact Summary
        st.subheader("Generate Cohort Impact Report")
        
        # Check how many have completed POST
        post_complete = [p for p in participants if p.get('post_completed')]
        
        st.write(f"**{len(post_complete)}** of **{len(participants)}** participants have completed POST assessments.")
        
        if len(post_complete) < 2:
            st.warning("Need at least 2 completed POST assessments to generate a cohort report.")
            return
        
        if st.button("Generate Impact Report", type="primary"):
            with st.spinner("Generating report (this may take a moment for AI theme analysis)..."):
                try:
                    doc_buffer = report_gen.generate_impact_report(selected_cohort_id)
                    cohort = db.get_cohort(selected_cohort_id)
                    st.download_button(
                        label="📥 Download Impact Report",
                        data=doc_buffer,
                        file_name=f"Readiness_Impact_{cohort['name'].replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.success("Report generated successfully!")
                except Exception as e:
                    st.error(f"Error generating report: {e}")


def show_settings():
    """Display settings and system info."""
    
    st.header("Settings")
    
    # Database info
    st.subheader("Database Connection")
    db_info = db.get_db_info()
    
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Type:** {db_info['type']}")
    with col2:
        st.write(f"**Status:** {db_info['status']}")
    
    if db_info['type'] == 'Turso Cloud':
        st.write(f"**URL:** {db_info['url']}")
    
    st.divider()
    
    # Theme extractor status
    st.subheader("AI Theme Extraction")
    
    from theme_extractor import ThemeExtractor
    extractor = ThemeExtractor()
    
    if extractor.is_available():
        st.success("✅ Claude API connected - AI theme extraction available")
    else:
        st.warning("⚠️ Claude API not configured - theme extraction will be limited")
        st.caption("Add your Anthropic API key to Streamlit secrets to enable AI-powered theme analysis.")
    
    st.divider()
    
    # Framework info
    st.subheader("Assessment Framework")
    
    st.write(f"**Indicators:** {len(INDICATORS)}")
    st.write(f"**Items:** {len(ITEMS)}")
    st.write(f"**Rating Scale:** 1-6")
    
    with st.expander("View Framework Details"):
        for indicator, (start, end) in INDICATORS.items():
            st.write(f"**{indicator}** (Items {start}-{end})")
            st.caption(INDICATOR_DESCRIPTIONS.get(indicator, ""))


def show_assessment_form(token: str):
    """Display the assessment form for participants."""
    
    # Import the assessment form module
    from assessment_form import show_assessment
    show_assessment(db, token)


if __name__ == "__main__":
    main()
