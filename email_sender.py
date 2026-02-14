#!/usr/bin/env python3
"""
Email module for Launch Readiness Assessment System.

Sends branded assessment invitations and reminders via SendGrid.
Tracks all sent emails in the database.
"""

import streamlit as st
from datetime import datetime

try:
    from sendgrid import SendGridAPIClient
    from sendgrid.helpers.mail import Mail, Email, To, Content, HtmlContent
    SENDGRID_AVAILABLE = True
except ImportError:
    SENDGRID_AVAILABLE = False


def get_sendgrid_client():
    """Get SendGrid client from Streamlit secrets."""
    try:
        api_key = st.secrets.get("sendgrid", {}).get("api_key", "")
        sender_email = st.secrets.get("sendgrid", {}).get("sender_email", "")
        sender_name = st.secrets.get("sendgrid", {}).get("sender_name", "The Development Catalyst")
        
        if api_key and sender_email:
            return SendGridAPIClient(api_key), sender_email, sender_name
    except Exception:
        pass
    return None, None, None


def is_email_configured():
    """Check if SendGrid is properly configured."""
    if not SENDGRID_AVAILABLE:
        return False
    client, email, name = get_sendgrid_client()
    return client is not None


def send_assessment_email(participant, assessment_type, base_url, db):
    """
    Send an assessment invitation email.
    
    Args:
        participant: dict with name, email, role, pre_token, post_token
        assessment_type: 'PRE' or 'POST'
        base_url: app base URL for building assessment links
        db: database instance for logging
    
    Returns:
        (success: bool, message: str)
    """
    if not participant.get('email'):
        return False, f"No email address for {participant['name']}"
    
    client, sender_email, sender_name = get_sendgrid_client()
    if not client:
        return False, "SendGrid not configured"
    
    # Build the assessment link
    token = participant['pre_token'] if assessment_type == 'PRE' else participant['post_token']
    assessment_url = f"{base_url}?token={token}"
    
    # Build the email
    first_name = participant['name'].split()[0]
    
    if assessment_type == 'PRE':
        subject = "Launch Readiness — Your Pre-Programme Assessment"
        html_content = _build_pre_email(first_name, participant['name'], assessment_url)
    else:
        subject = "Launch Readiness — Your Post-Programme Assessment"
        html_content = _build_post_email(first_name, participant['name'], assessment_url)
    
    message = Mail(
        from_email=Email(sender_email, sender_name),
        to_emails=To(participant['email'], participant['name']),
        subject=subject,
        html_content=HtmlContent(html_content)
    )
    
    try:
        response = client.send(message)
        
        # Log the send
        db.log_email(
            participant_id=participant['id'],
            email_type=assessment_type,
            recipient_email=participant['email'],
            status='sent' if response.status_code in (200, 201, 202) else 'failed',
            status_code=response.status_code
        )
        
        if response.status_code in (200, 201, 202):
            return True, f"Sent to {participant['email']}"
        else:
            return False, f"SendGrid returned status {response.status_code}"
    
    except Exception as e:
        db.log_email(
            participant_id=participant['id'],
            email_type=assessment_type,
            recipient_email=participant['email'],
            status='error',
            status_code=0,
            error_message=str(e)
        )
        return False, f"Error: {e}"


def send_reminder_email(participant, assessment_type, base_url, db):
    """Send a reminder email for an incomplete assessment."""
    if not participant.get('email'):
        return False, f"No email address for {participant['name']}"
    
    client, sender_email, sender_name = get_sendgrid_client()
    if not client:
        return False, "SendGrid not configured"
    
    token = participant['pre_token'] if assessment_type == 'PRE' else participant['post_token']
    assessment_url = f"{base_url}?token={token}"
    first_name = participant['name'].split()[0]
    
    if assessment_type == 'PRE':
        subject = "Reminder — Your Pre-Programme Assessment"
    else:
        subject = "Reminder — Your Post-Programme Assessment"
    
    html_content = _build_reminder_email(first_name, participant['name'], assessment_url, assessment_type)
    
    message = Mail(
        from_email=Email(sender_email, sender_name),
        to_emails=To(participant['email'], participant['name']),
        subject=subject,
        html_content=HtmlContent(html_content)
    )
    
    try:
        response = client.send(message)
        
        db.log_email(
            participant_id=participant['id'],
            email_type=f'{assessment_type}_REMINDER',
            recipient_email=participant['email'],
            status='sent' if response.status_code in (200, 201, 202) else 'failed',
            status_code=response.status_code
        )
        
        if response.status_code in (200, 201, 202):
            return True, f"Reminder sent to {participant['email']}"
        else:
            return False, f"SendGrid returned status {response.status_code}"
    
    except Exception as e:
        db.log_email(
            participant_id=participant['id'],
            email_type=f'{assessment_type}_REMINDER',
            recipient_email=participant['email'],
            status='error',
            status_code=0,
            error_message=str(e)
        )
        return False, f"Error: {e}"


# =========== EMAIL TEMPLATES ===========

def _build_pre_email(first_name, full_name, assessment_url):
    """Build branded PRE assessment invitation email."""
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin:0; padding:0; background-color:#F5F5F5; font-family:Arial, sans-serif;">
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#F5F5F5; padding:20px 0;">
            <tr>
                <td align="center">
                    <table width="600" cellpadding="0" cellspacing="0" style="background-color:#FFFFFF; border-radius:8px; overflow:hidden;">
                        
                        <!-- Header -->
                        <tr>
                            <td style="background-color:#461E96; padding:30px 40px; text-align:center;">
                                <h1 style="color:#FFFFFF; margin:0; font-size:22px; font-weight:bold;">THE READINESS FRAMEWORK</h1>
                                <p style="color:#E6008C; margin:8px 0 0; font-size:14px;">Pre-Programme Assessment</p>
                            </td>
                        </tr>
                        
                        <!-- Body -->
                        <tr>
                            <td style="padding:30px 40px;">
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    Hi {first_name},
                                </p>
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    Welcome to the Launch Readiness programme. Before we begin, we'd like you to complete 
                                    a short self-assessment. This takes about 10 minutes and helps us understand where 
                                    you see yourself today.
                                </p>
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    There are no right or wrong answers — this is simply a snapshot of your starting point. 
                                    Your responses are confidential and will be used to create your personal development report.
                                </p>
                                
                                <!-- CTA Button -->
                                <table width="100%" cellpadding="0" cellspacing="0" style="margin:25px 0;">
                                    <tr>
                                        <td align="center">
                                            <a href="{assessment_url}" 
                                               style="background-color:#461E96; color:#FFFFFF; padding:14px 40px; 
                                                      text-decoration:none; border-radius:6px; font-size:16px; 
                                                      font-weight:bold; display:inline-block;">
                                                Complete Your Assessment
                                            </a>
                                        </td>
                                    </tr>
                                </table>
                                
                                <p style="color:#6E6E6E; font-size:13px; line-height:1.5; margin:15px 0 0;">
                                    If the button doesn't work, copy and paste this link into your browser:<br>
                                    <a href="{assessment_url}" style="color:#461E96; word-break:break-all;">{assessment_url}</a>
                                </p>
                            </td>
                        </tr>
                        
                        <!-- Footer -->
                        <tr>
                            <td style="background-color:#F5F5F5; padding:20px 40px; text-align:center; border-top:1px solid #E8E8E8;">
                                <p style="color:#6E6E6E; font-size:11px; margin:0;">
                                    This assessment is delivered by The Development Catalyst on behalf of Cencora.
                                </p>
                                <p style="color:#6E6E6E; font-size:11px; margin:5px 0 0;">
                                    Your responses are confidential.
                                </p>
                            </td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """


def _build_post_email(first_name, full_name, assessment_url):
    """Build branded POST assessment invitation email."""
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin:0; padding:0; background-color:#F5F5F5; font-family:Arial, sans-serif;">
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#F5F5F5; padding:20px 0;">
            <tr>
                <td align="center">
                    <table width="600" cellpadding="0" cellspacing="0" style="background-color:#FFFFFF; border-radius:8px; overflow:hidden;">
                        
                        <!-- Header -->
                        <tr>
                            <td style="background-color:#461E96; padding:30px 40px; text-align:center;">
                                <h1 style="color:#FFFFFF; margin:0; font-size:22px; font-weight:bold;">THE READINESS FRAMEWORK</h1>
                                <p style="color:#00DC8C; margin:8px 0 0; font-size:14px;">Post-Programme Assessment</p>
                            </td>
                        </tr>
                        
                        <!-- Body -->
                        <tr>
                            <td style="padding:30px 40px;">
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    Hi {first_name},
                                </p>
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    Congratulations on completing the Launch Readiness programme! We'd now like you to 
                                    complete the same assessment again so we can measure your growth.
                                </p>
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    This takes about 10 minutes. You'll also have the chance to reflect on your biggest 
                                    takeaways and what you'll commit to doing differently. Your responses will be used 
                                    to create your personal progress report.
                                </p>
                                
                                <!-- CTA Button -->
                                <table width="100%" cellpadding="0" cellspacing="0" style="margin:25px 0;">
                                    <tr>
                                        <td align="center">
                                            <a href="{assessment_url}" 
                                               style="background-color:#007F50; color:#FFFFFF; padding:14px 40px; 
                                                      text-decoration:none; border-radius:6px; font-size:16px; 
                                                      font-weight:bold; display:inline-block;">
                                                Complete Your Assessment
                                            </a>
                                        </td>
                                    </tr>
                                </table>
                                
                                <p style="color:#6E6E6E; font-size:13px; line-height:1.5; margin:15px 0 0;">
                                    If the button doesn't work, copy and paste this link into your browser:<br>
                                    <a href="{assessment_url}" style="color:#461E96; word-break:break-all;">{assessment_url}</a>
                                </p>
                            </td>
                        </tr>
                        
                        <!-- Footer -->
                        <tr>
                            <td style="background-color:#F5F5F5; padding:20px 40px; text-align:center; border-top:1px solid #E8E8E8;">
                                <p style="color:#6E6E6E; font-size:11px; margin:0;">
                                    This assessment is delivered by The Development Catalyst on behalf of Cencora.
                                </p>
                                <p style="color:#6E6E6E; font-size:11px; margin:5px 0 0;">
                                    Your responses are confidential.
                                </p>
                            </td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """


def _build_reminder_email(first_name, full_name, assessment_url, assessment_type):
    """Build a reminder email."""
    stage = "pre-programme" if assessment_type == "PRE" else "post-programme"
    
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin:0; padding:0; background-color:#F5F5F5; font-family:Arial, sans-serif;">
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#F5F5F5; padding:20px 0;">
            <tr>
                <td align="center">
                    <table width="600" cellpadding="0" cellspacing="0" style="background-color:#FFFFFF; border-radius:8px; overflow:hidden;">
                        
                        <!-- Header -->
                        <tr>
                            <td style="background-color:#461E96; padding:30px 40px; text-align:center;">
                                <h1 style="color:#FFFFFF; margin:0; font-size:22px; font-weight:bold;">THE READINESS FRAMEWORK</h1>
                                <p style="color:#FFA400; margin:8px 0 0; font-size:14px;">Friendly Reminder</p>
                            </td>
                        </tr>
                        
                        <!-- Body -->
                        <tr>
                            <td style="padding:30px 40px;">
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    Hi {first_name},
                                </p>
                                <p style="color:#3B3B3B; font-size:15px; line-height:1.6; margin:0 0 15px;">
                                    Just a quick reminder that your {stage} assessment is still waiting for you. 
                                    It only takes about 10 minutes and your input is really valuable.
                                </p>
                                
                                <!-- CTA Button -->
                                <table width="100%" cellpadding="0" cellspacing="0" style="margin:25px 0;">
                                    <tr>
                                        <td align="center">
                                            <a href="{assessment_url}" 
                                               style="background-color:#461E96; color:#FFFFFF; padding:14px 40px; 
                                                      text-decoration:none; border-radius:6px; font-size:16px; 
                                                      font-weight:bold; display:inline-block;">
                                                Complete Your Assessment
                                            </a>
                                        </td>
                                    </tr>
                                </table>
                                
                                <p style="color:#6E6E6E; font-size:13px; line-height:1.5; margin:15px 0 0;">
                                    If the button doesn't work, copy and paste this link into your browser:<br>
                                    <a href="{assessment_url}" style="color:#461E96; word-break:break-all;">{assessment_url}</a>
                                </p>
                            </td>
                        </tr>
                        
                        <!-- Footer -->
                        <tr>
                            <td style="background-color:#F5F5F5; padding:20px 40px; text-align:center; border-top:1px solid #E8E8E8;">
                                <p style="color:#6E6E6E; font-size:11px; margin:0;">
                                    This assessment is delivered by The Development Catalyst on behalf of Cencora.
                                </p>
                                <p style="color:#6E6E6E; font-size:11px; margin:5px 0 0;">
                                    Your responses are confidential.
                                </p>
                            </td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """
