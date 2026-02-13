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
    
    # Page styling
    st.markdown("""
    <style>
        .assessment-header {
            color: #461E96;
            font-size: 1.8rem;
            font-weight: bold;
            margin-bottom: 0.5rem;
        }
        .assessment-subheader {
            color: #E6008C;
            font-size: 1.1rem;
            margin-bottom: 1.5rem;
        }
        .indicator-header {
            font-size: 1.2rem;
            font-weight: bold;
            margin-top: 1.5rem;
            margin-bottom: 0.5rem;
            padding: 0.5rem;
            border-radius: 4px;
            color: white;
        }
        .welcome-box {
            background-color: #F5F5F5;
            padding: 1.5rem;
            border-radius: 8px;
            border-left: 4px solid #461E96;
            margin-bottom: 1.5rem;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    assessment_title = "Pre-Programme Assessment" if is_pre else "Post-Programme Assessment"
    st.markdown(f'<p class="assessment-header">🚀 Launch Readiness</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="assessment-subheader">{assessment_title}</p>', unsafe_allow_html=True)
    
    # Welcome message
    if is_pre:
        welcome_text = f"""
        Welcome, **{participant['name']}**! 
        
        This assessment will help you reflect on your current readiness as you prepare for your new role.
        There are no right or wrong answers - this is simply a snapshot of where you see yourself today.
        
        The assessment takes about **10-15 minutes** to complete. Please answer honestly - your responses
        are confidential and will be used to support your development.
        """
    else:
        welcome_text = f"""
        Welcome back, **{participant['name']}**!
        
        Now that you've completed the Launch Readiness programme, this assessment will help capture
        your growth and identify areas for continued development.
        
        As before, there are no right or wrong answers. Your honest reflection will help us understand
        the programme's impact and support your ongoing development.
        """
    
    st.markdown(f'<div class="welcome-box">{welcome_text}</div>', unsafe_allow_html=True)
    
    st.divider()
    
    # Assessment form
    with st.form("assessment_form"):
        ratings = {}
        
        # Rating scale reminder
        st.caption("**Rating Scale:** 1 = Strongly Disagree, 2 = Disagree, 3 = Slightly Disagree, 4 = Slightly Agree, 5 = Agree, 6 = Strongly Agree")
        
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
                
                col1, col2 = st.columns([4, 2])
                with col1:
                    st.write(f"**{item_num}.** {item['text']}")
                with col2:
                    ratings[item_num] = st.select_slider(
                        f"Rating for item {item_num}",
                        options=[1, 2, 3, 4, 5, 6],
                        value=4,
                        key=f"rating_{item_num}",
                        label_visibility="collapsed"
                    )
            
            st.write("")  # Spacing
        
        # Overall Readiness
        st.markdown(
            '<div class="indicator-header" style="background-color: #461E96;">Overall Readiness</div>',
            unsafe_allow_html=True
        )
        
        for item_num in [31, 32]:
            item = ITEMS[item_num]
            col1, col2 = st.columns([4, 2])
            with col1:
                st.write(f"**{item_num}.** {item['text']}")
            with col2:
                ratings[item_num] = st.select_slider(
                    f"Rating for item {item_num}",
                    options=[1, 2, 3, 4, 5, 6],
                    value=4,
                    key=f"rating_{item_num}",
                    label_visibility="collapsed"
                )
        
        st.divider()
        
        # Open questions
        st.subheader("Your Reflections")
        
        open_responses = {}
        
        for q_num, question in open_questions.items():
            # Special handling for POST question 3 (show original concern)
            if not is_pre and q_num == 3 and pre_concern:
                st.write(f"**{q_num}. {question}**")
                st.info(f"**Your original concern:** \"{pre_concern}\"")
                open_responses[q_num] = st.text_area(
                    "Your reflection",
                    key=f"open_{q_num}",
                    height=100,
                    placeholder="How do you feel about this now?",
                    label_visibility="collapsed"
                )
            else:
                open_responses[q_num] = st.text_area(
                    f"**{q_num}. {question}**",
                    key=f"open_{q_num}",
                    height=100,
                    placeholder="Share your thoughts..."
                )
        
        st.divider()
        
        # Submit button
        submitted = st.form_submit_button("Submit Assessment", type="primary", use_container_width=True)
        
        if submitted:
            # Validate all ratings are provided (they have defaults, so should be fine)
            if len(ratings) < 32:
                st.error("Please complete all rating questions.")
            else:
                # Save ratings
                db.save_all_ratings(assessment['id'], ratings)
                
                # Save open responses
                db.save_all_open_responses(assessment['id'], open_responses)
                
                # Mark as completed
                db.mark_assessment_completed(token)
                
                st.success("Thank you! Your assessment has been submitted successfully.")
                st.balloons()
                
                # Show completion message
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
        }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<p class="assessment-header">🚀 Launch Readiness</p>', unsafe_allow_html=True)
    
    if assessment_type == 'PRE':
        message = """
        <div class="completion-box">
            <h2>✅ Assessment Complete</h2>
            <p>You have already completed your <strong>Pre-Programme Assessment</strong>.</p>
            <p>Thank you for your responses. We look forward to seeing you in the programme!</p>
            <p><em>You will receive a link to complete your Post-Programme Assessment after the programme ends.</em></p>
        </div>
        """
    else:
        message = """
        <div class="completion-box">
            <h2>✅ Assessment Complete</h2>
            <p>You have already completed your <strong>Post-Programme Assessment</strong>.</p>
            <p>Thank you for your participation in the Launch Readiness programme!</p>
            <p><em>Your facilitator will share your Progress Report with you shortly.</em></p>
        </div>
        """
    
    st.markdown(message, unsafe_allow_html=True)
