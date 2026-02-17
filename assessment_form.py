#!/usr/bin/env python3
"""
Assessment form for participants.

Displays PRE or POST assessment based on the access token.
"""

import streamlit as st
from framework import (
    INDICATORS, INDICATOR_DESCRIPTIONS, INDICATOR_COLOURS,
    ITEMS, OPEN_QUESTIONS_PRE, OPEN_QUESTIONS_POST, RATING_SCALE,
    get_items_for_indicator
)


def show_assessment(db, token: str):
    """Display the assessment form."""
    
    # Validate token
    assessment = db.get_assessment_by_token(token)
    
    if not assessment:
        st.error("Invalid or expired assessment link.")
        st.stop()
    
    # Check if already completed
    if assessment.get('completed_at'):
        show_completion_message(assessment['assessment_type'])
        return
    
    # Get participant info
    participant = db.get_participant(assessment['participant_id'])
    cohort = db.get_cohort(participant['cohort_id'])
    
    # Mark as started
    db.mark_assessment_started(token)
    
    # Determine assessment type
    is_pre = assessment['assessment_type'] == 'PRE'
    open_questions = OPEN_QUESTIONS_PRE if is_pre else OPEN_QUESTIONS_POST
    
    # For POST assessment, get PRE responses to show concern reflection
    pre_concern = None
    if not is_pre:
        assessments = db.get_assessments_for_participant(participant['id'])
        if assessments['PRE'] and assessments['PRE'].get('completed_at'):
            pre_responses = db.get_open_responses(assessments['PRE']['id'])
            pre_concern = pre_responses.get(3, "")  # Question 3 was concerns
    
    # Page styling - mobile friendly
    st.markdown("""
    <style>
        .main-title {
            color: #461E96;
            font-size: 1.8rem;
            font-weight: bold;
            margin-bottom: 0.5rem;
        }
        .sub-title {
            color: #E6008C;
            font-size: 1.1rem;
            margin-bottom: 1.5rem;
        }
        .indicator-header {
            font-size: 1.1rem;
            font-weight: bold;
            margin-top: 1.5rem;
            margin-bottom: 0.5rem;
            padding: 0.75rem;
            border-radius: 4px;
            color: white;
        }
        .welcome-box {
            background-color: #F5F5F5;
            padding: 1.2rem;
            border-radius: 8px;
            border-left: 4px solid #461E96;
            margin-bottom: 1.5rem;
            line-height: 1.6;
        }
        .welcome-box strong {
            color: #461E96;
        }
        .question-text {
            font-size: 0.95rem;
            line-height: 1.4;
            margin-bottom: 0.5rem;
        }
        .item-box {
            background-color: #F5F5F5;
            padding: 0.75rem 1rem;
            border-radius: 6px;
            border-left: 3px solid #461E96;
            margin-bottom: 0.25rem;
        }
        .warning-box {
            background-color: #FFF3CD;
            border: 1px solid #FFCC00;
            border-radius: 8px;
            padding: 1rem;
            margin: 1rem 0;
        }
        /* Mobile responsive */
        @media (max-width: 768px) {
            .main-title { font-size: 1.4rem; }
            .indicator-header { font-size: 1rem; padding: 0.5rem; }
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    assessment_title = "Pre-Programme Assessment" if is_pre else "Post-Programme Assessment"
    st.markdown(f'<p class="main-title">🚀 Launch Readiness</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-title">{assessment_title}</p>', unsafe_allow_html=True)
    
    # Welcome message - using proper HTML
    first_name = participant['name'].split()[0]
    if is_pre:
        welcome_html = f"""
        <div class="welcome-box">
            <p>Welcome, <strong>{first_name}</strong>!</p>
            <p>This assessment will help you reflect on your current readiness as you prepare for your new role.
            There are no right or wrong answers – this is simply a snapshot of where you see yourself today.</p>
            <p>The assessment takes about <strong>10-15 minutes</strong> to complete. Please answer honestly – your responses
            are confidential and will be used to support your development.</p>
        </div>
        """
    else:
        welcome_html = f"""
        <div class="welcome-box">
            <p>Welcome back, <strong>{first_name}</strong>!</p>
            <p>Now that you've completed the Launch Readiness programme, this assessment will help capture
            your growth and identify areas for continued development.</p>
            <p>As before, there are no right or wrong answers. Your honest reflection will help us understand
            the programme's impact and support your ongoing development.</p>
        </div>
        """
    
    st.markdown(welcome_html, unsafe_allow_html=True)
    
    st.divider()
    
    # Assessment form
    with st.form("assessment_form"):
        ratings = {}
        
        # Rating scale reminder
        st.markdown("**Rating Scale:**")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.caption("1 = Strongly Disagree")
            st.caption("2 = Disagree")
        with col2:
            st.caption("3 = Slightly Disagree")
            st.caption("4 = Slightly Agree")
        with col3:
            st.caption("5 = Agree")
            st.caption("6 = Strongly Agree")
        
        st.divider()
        
        # Loop through indicators
        for indicator, (start, end) in INDICATORS.items():
            colour = INDICATOR_COLOURS.get(indicator, '#461E96')
            st.markdown(
                f'<div class="indicator-header" style="background-color: {colour};">{indicator}</div>',
                unsafe_allow_html=True
            )
            st.caption(INDICATOR_DESCRIPTIONS.get(indicator, ""))
            
            # Items for this indicator
            for item_num in range(start, end + 1):
                item = ITEMS[item_num]
                
                st.markdown(f'<div class="item-box" style="border-left-color: {colour};"><strong>{item_num}.</strong> {item["text"]}</div>', unsafe_allow_html=True)
                
                ratings[item_num] = st.radio(
                    f"Rating for item {item_num}",
                    options=[1, 2, 3, 4, 5, 6],
                    format_func=lambda x: {1: "1 – Strongly Disagree", 2: "2 – Disagree", 3: "3 – Slightly Disagree", 4: "4 – Slightly Agree", 5: "5 – Agree", 6: "6 – Strongly Agree"}[x],
                    index=None,
                    key=f"rating_{item_num}",
                    label_visibility="collapsed",
                    horizontal=True
                )
                
                st.write("")  # Spacing between questions
            
            st.divider()
        
        # Overall Readiness
        st.markdown(
            '<div class="indicator-header" style="background-color: #461E96;">Overall Readiness</div>',
            unsafe_allow_html=True
        )
        
        for item_num in [31, 32]:
            item = ITEMS[item_num]
            st.markdown(f'<div class="item-box"><strong>{item_num}.</strong> {item["text"]}</div>', unsafe_allow_html=True)
            ratings[item_num] = st.radio(
                f"Rating for item {item_num}",
                options=[1, 2, 3, 4, 5, 6],
                format_func=lambda x: {1: "1 – Strongly Disagree", 2: "2 – Disagree", 3: "3 – Slightly Disagree", 4: "4 – Slightly Agree", 5: "5 – Agree", 6: "6 – Strongly Agree"}[x],
                index=None,
                key=f"rating_{item_num}",
                label_visibility="collapsed",
                horizontal=True
            )
            st.write("")
        
        st.divider()
        
        # Open questions
        st.markdown('<div class="indicator-header" style="background-color: #461E96;">Your Reflections</div>', unsafe_allow_html=True)
        
        open_responses = {}
        
        for q_num, question in open_questions.items():
            # Special handling for POST question 3 (show original concern)
            if not is_pre and q_num == 3 and pre_concern:
                st.markdown(f"**{q_num}. {question}**")
                st.info(f"**Your original concern:** \"{pre_concern}\"")
                open_responses[q_num] = st.text_area(
                    "Your reflection",
                    key=f"open_{q_num}",
                    height=100,
                    placeholder="How do you feel about this now?",
                    label_visibility="collapsed"
                )
            else:
                st.markdown(f"**{q_num}. {question}**")
                open_responses[q_num] = st.text_area(
                    f"Response to question {q_num}",
                    key=f"open_{q_num}",
                    height=100,
                    placeholder="Share your thoughts...",
                    label_visibility="collapsed"
                )
            st.write("")
        
        st.divider()
        
        # Submit button
        submitted = st.form_submit_button("Submit Assessment", type="primary", use_container_width=True)
        
        if submitted:
            # Check for unanswered items (None = no selection made)
            unanswered_items = [num for num, score in ratings.items() if score is None]
            
            if unanswered_items:
                item_list = ""
                for num in unanswered_items:
                    item_list += f"  - **Item {num}:** {ITEMS[num]['text']}\n"
                
                st.error(f"""
⚠️ **Please complete all items before submitting**

You have **{len(unanswered_items)}** unanswered item{'s' if len(unanswered_items) > 1 else ''}:

{item_list}

Please scroll up and select a rating for each one.
                """)
            else:
                # Save ratings
                db.save_all_ratings(assessment['id'], ratings)
                
                # Save open responses
                db.save_all_open_responses(assessment['id'], open_responses)
                
                # Mark as completed
                db.mark_assessment_completed(token)
                
                st.success("Thank you! Your assessment has been submitted successfully.")
                st.balloons()
                
                # Rerun to show completion message
                st.rerun()


def show_completion_message(assessment_type: str):
    """Show message for already completed assessment."""
    
    st.markdown("""
    <style>
        .completion-box {
            background-color: #E8F5E9;
            padding: 2rem;
            border-radius: 8px;
            text-align: center;
            border: 2px solid #00DC8C;
            margin-top: 2rem;
        }
        .completion-box h2 {
            color: #2E7D32;
            margin-bottom: 1rem;
        }
        .completion-box p {
            color: #3B3B3B;
            margin-bottom: 0.5rem;
        }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<p class="main-title">🚀 Launch Readiness</p>', unsafe_allow_html=True)
    
    if assessment_type == 'PRE':
        st.markdown("""
        <div class="completion-box">
            <h2>✅ Assessment Complete</h2>
            <p>You have already completed your <strong>Pre-Programme Assessment</strong>.</p>
            <p>Thank you for your responses. We look forward to seeing you in the programme!</p>
            <p><em>You will receive a link to complete your Post-Programme Assessment after the programme ends.</em></p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="completion-box">
            <h2>✅ Assessment Complete</h2>
            <p>You have already completed your <strong>Post-Programme Assessment</strong>.</p>
            <p>Thank you for your participation in the Launch Readiness programme!</p>
            <p><em>Your facilitator will share your Progress Report with you shortly.</em></p>
        </div>
        """, unsafe_allow_html=True)
