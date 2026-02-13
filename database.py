#!/usr/bin/env python3
"""
Database module for the Cencora Launch Readiness system.

Handles all data persistence using Turso (libSQL).
"""

import sqlite3
import secrets
import hashlib
from datetime import datetime
import os

# Try to import libsql for Turso cloud connection
try:
    import libsql_experimental as libsql
    USING_TURSO = True
except ImportError:
    USING_TURSO = False


class Database:
    def __init__(self, db_path="readiness.db"):
        self.db_path = db_path
        self.turso_url = None
        self.turso_token = None
        
        # Try Streamlit secrets first
        try:
            import streamlit as st
            self.turso_url = st.secrets.get("turso", {}).get("url")
            self.turso_token = st.secrets.get("turso", {}).get("token")
        except:
            pass
        
        # Fall back to environment variables
        if not self.turso_url:
            self.turso_url = os.environ.get("TURSO_DATABASE_URL")
        if not self.turso_token:
            self.turso_token = os.environ.get("TURSO_AUTH_TOKEN")
        
        self.init_database()
    
    def get_connection(self):
        """Get a database connection."""
        if self.turso_url and self.turso_token and USING_TURSO:
            # Connect to Turso cloud database
            conn = libsql.connect(
                database=self.turso_url,
                auth_token=self.turso_token
            )
            return conn
        else:
            # Fall back to local SQLite
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            return conn
    
    def _execute(self, query, params=None):
        """Execute a query and return the cursor."""
        conn = self.get_connection()
        cursor = conn.cursor()
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)
        return conn, cursor
    
    def _fetchall(self, query, params=None):
        """Execute a query and fetch all results as list of dicts."""
        conn, cursor = self._execute(query, params)
        
        if USING_TURSO and self.turso_url and self.turso_token:
            rows = cursor.fetchall()
            if rows and len(rows) > 0:
                columns = [desc[0] for desc in cursor.description]
                result = [dict(zip(columns, row)) for row in rows]
            else:
                result = []
        else:
            result = [dict(row) for row in cursor.fetchall()]
        
        conn.close()
        return result
    
    def _fetchone(self, query, params=None):
        """Execute a query and fetch one result as dict."""
        conn, cursor = self._execute(query, params)
        row = cursor.fetchone()
        
        if row:
            if USING_TURSO and self.turso_url and self.turso_token:
                columns = [desc[0] for desc in cursor.description]
                result = dict(zip(columns, row))
            else:
                result = dict(row)
        else:
            result = None
        
        conn.close()
        return result
    
    def init_database(self):
        """Initialize the database schema."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Cohorts table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cohorts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                programme TEXT DEFAULT 'Launch Readiness',
                description TEXT,
                start_date TEXT,
                end_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Participants table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS participants (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cohort_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                email TEXT,
                role TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (cohort_id) REFERENCES cohorts (id)
            )
        ''')
        
        # Assessments table (PRE and POST)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS assessments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                participant_id INTEGER NOT NULL,
                assessment_type TEXT NOT NULL CHECK (assessment_type IN ('PRE', 'POST')),
                access_token TEXT UNIQUE NOT NULL,
                started_at TEXT,
                completed_at TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (participant_id) REFERENCES participants (id)
            )
        ''')
        
        # Ratings table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS ratings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                assessment_id INTEGER NOT NULL,
                item_number INTEGER NOT NULL,
                score INTEGER NOT NULL CHECK (score >= 1 AND score <= 6),
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (assessment_id) REFERENCES assessments (id),
                UNIQUE (assessment_id, item_number)
            )
        ''')
        
        # Open responses table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS open_responses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                assessment_id INTEGER NOT NULL,
                question_number INTEGER NOT NULL,
                response_text TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (assessment_id) REFERENCES assessments (id),
                UNIQUE (assessment_id, question_number)
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def get_db_info(self):
        """Get database connection info."""
        if self.turso_url and self.turso_token and USING_TURSO:
            return {
                'type': 'Turso Cloud',
                'url': self.turso_url,
                'status': 'Connected'
            }
        else:
            return {
                'type': 'Local SQLite',
                'path': self.db_path,
                'status': 'Connected'
            }
    
    # =========== COHORT OPERATIONS ===========
    
    def create_cohort(self, name: str, programme: str = "Launch Readiness", 
                      description: str = None, start_date: str = None, end_date: str = None) -> int:
        """Create a new cohort and return its ID."""
        conn, cursor = self._execute(
            '''INSERT INTO cohorts (name, programme, description, start_date, end_date)
               VALUES (?, ?, ?, ?, ?)''',
            (name, programme, description, start_date, end_date)
        )
        cohort_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return cohort_id
    
    def get_cohort(self, cohort_id: int) -> dict:
        """Get a cohort by ID."""
        return self._fetchone('SELECT * FROM cohorts WHERE id = ?', (cohort_id,))
    
    def get_all_cohorts(self) -> list:
        """Get all cohorts."""
        return self._fetchall('SELECT * FROM cohorts ORDER BY created_at DESC')
    
    def update_cohort(self, cohort_id: int, **kwargs):
        """Update cohort fields."""
        valid_fields = ['name', 'programme', 'description', 'start_date', 'end_date']
        updates = {k: v for k, v in kwargs.items() if k in valid_fields}
        
        if not updates:
            return
        
        set_clause = ', '.join(f'{k} = ?' for k in updates.keys())
        values = list(updates.values()) + [cohort_id]
        
        conn, cursor = self._execute(
            f'UPDATE cohorts SET {set_clause} WHERE id = ?', values
        )
        conn.commit()
        conn.close()
    
    def delete_cohort(self, cohort_id: int):
        """Delete a cohort and all related data."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Get all participants in cohort
        cursor.execute('SELECT id FROM participants WHERE cohort_id = ?', (cohort_id,))
        participants = cursor.fetchall()
        
        for p in participants:
            p_id = p[0] if isinstance(p, tuple) else p['id']
            # Get all assessments for participant
            cursor.execute('SELECT id FROM assessments WHERE participant_id = ?', (p_id,))
            assessments = cursor.fetchall()
            
            for a in assessments:
                a_id = a[0] if isinstance(a, tuple) else a['id']
                cursor.execute('DELETE FROM ratings WHERE assessment_id = ?', (a_id,))
                cursor.execute('DELETE FROM open_responses WHERE assessment_id = ?', (a_id,))
            
            cursor.execute('DELETE FROM assessments WHERE participant_id = ?', (p_id,))
        
        cursor.execute('DELETE FROM participants WHERE cohort_id = ?', (cohort_id,))
        cursor.execute('DELETE FROM cohorts WHERE id = ?', (cohort_id,))
        
        conn.commit()
        conn.close()
    
    # =========== PARTICIPANT OPERATIONS ===========
    
    def create_participant(self, cohort_id: int, name: str, email: str = None, role: str = None) -> int:
        """Create a new participant and return their ID."""
        conn, cursor = self._execute(
            '''INSERT INTO participants (cohort_id, name, email, role)
               VALUES (?, ?, ?, ?)''',
            (cohort_id, name, email, role)
        )
        participant_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        # Automatically create PRE and POST assessments
        self.create_assessment(participant_id, 'PRE')
        self.create_assessment(participant_id, 'POST')
        
        return participant_id
    
    def get_participant(self, participant_id: int) -> dict:
        """Get a participant by ID."""
        return self._fetchone('SELECT * FROM participants WHERE id = ?', (participant_id,))
    
    def get_participant_by_token(self, token: str) -> dict:
        """Get participant info from an assessment token."""
        assessment = self._fetchone(
            'SELECT participant_id, assessment_type FROM assessments WHERE access_token = ?',
            (token,)
        )
        if not assessment:
            return None
        
        participant = self.get_participant(assessment['participant_id'])
        if participant:
            participant['assessment_type'] = assessment['assessment_type']
        return participant
    
    def get_participants_for_cohort(self, cohort_id: int) -> list:
        """Get all participants in a cohort with their assessment status."""
        participants = self._fetchall(
            'SELECT * FROM participants WHERE cohort_id = ? ORDER BY name',
            (cohort_id,)
        )
        
        for p in participants:
            # Get PRE assessment status
            pre = self._fetchone(
                '''SELECT access_token, completed_at FROM assessments 
                   WHERE participant_id = ? AND assessment_type = 'PRE' ''',
                (p['id'],)
            )
            p['pre_token'] = pre['access_token'] if pre else None
            p['pre_completed'] = pre['completed_at'] if pre else None
            
            # Get POST assessment status
            post = self._fetchone(
                '''SELECT access_token, completed_at FROM assessments 
                   WHERE participant_id = ? AND assessment_type = 'POST' ''',
                (p['id'],)
            )
            p['post_token'] = post['access_token'] if post else None
            p['post_completed'] = post['completed_at'] if post else None
        
        return participants
    
    def delete_participant(self, participant_id: int):
        """Delete a participant and all their data."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT id FROM assessments WHERE participant_id = ?', (participant_id,))
        assessments = cursor.fetchall()
        
        for a in assessments:
            a_id = a[0] if isinstance(a, tuple) else a['id']
            cursor.execute('DELETE FROM ratings WHERE assessment_id = ?', (a_id,))
            cursor.execute('DELETE FROM open_responses WHERE assessment_id = ?', (a_id,))
        
        cursor.execute('DELETE FROM assessments WHERE participant_id = ?', (participant_id,))
        cursor.execute('DELETE FROM participants WHERE id = ?', (participant_id,))
        
        conn.commit()
        conn.close()
    
    # =========== ASSESSMENT OPERATIONS ===========
    
    def _generate_token(self) -> str:
        """Generate a unique access token."""
        return secrets.token_urlsafe(32)
    
    def create_assessment(self, participant_id: int, assessment_type: str) -> str:
        """Create an assessment and return the access token."""
        token = self._generate_token()
        
        conn, cursor = self._execute(
            '''INSERT INTO assessments (participant_id, assessment_type, access_token)
               VALUES (?, ?, ?)''',
            (participant_id, assessment_type, token)
        )
        conn.commit()
        conn.close()
        return token
    
    def get_assessment(self, assessment_id: int) -> dict:
        """Get an assessment by ID."""
        return self._fetchone('SELECT * FROM assessments WHERE id = ?', (assessment_id,))
    
    def get_assessment_by_token(self, token: str) -> dict:
        """Get an assessment by access token."""
        return self._fetchone('SELECT * FROM assessments WHERE access_token = ?', (token,))
    
    def get_assessments_for_participant(self, participant_id: int) -> dict:
        """Get both PRE and POST assessments for a participant."""
        assessments = self._fetchall(
            'SELECT * FROM assessments WHERE participant_id = ?',
            (participant_id,)
        )
        result = {'PRE': None, 'POST': None}
        for a in assessments:
            result[a['assessment_type']] = a
        return result
    
    def mark_assessment_started(self, token: str):
        """Mark an assessment as started."""
        conn, cursor = self._execute(
            '''UPDATE assessments SET started_at = ? WHERE access_token = ? AND started_at IS NULL''',
            (datetime.now().isoformat(), token)
        )
        conn.commit()
        conn.close()
    
    def mark_assessment_completed(self, token: str):
        """Mark an assessment as completed."""
        conn, cursor = self._execute(
            '''UPDATE assessments SET completed_at = ? WHERE access_token = ?''',
            (datetime.now().isoformat(), token)
        )
        conn.commit()
        conn.close()
    
    def is_assessment_completed(self, token: str) -> bool:
        """Check if an assessment is already completed."""
        assessment = self.get_assessment_by_token(token)
        return assessment and assessment.get('completed_at') is not None
    
    # =========== RATING OPERATIONS ===========
    
    def save_rating(self, assessment_id: int, item_number: int, score: int):
        """Save or update a rating."""
        conn, cursor = self._execute(
            '''INSERT INTO ratings (assessment_id, item_number, score)
               VALUES (?, ?, ?)
               ON CONFLICT (assessment_id, item_number) 
               DO UPDATE SET score = ?''',
            (assessment_id, item_number, score, score)
        )
        conn.commit()
        conn.close()
    
    def save_all_ratings(self, assessment_id: int, ratings: dict):
        """Save all ratings at once. ratings = {item_number: score}"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        for item_number, score in ratings.items():
            cursor.execute(
                '''INSERT INTO ratings (assessment_id, item_number, score)
                   VALUES (?, ?, ?)
                   ON CONFLICT (assessment_id, item_number) 
                   DO UPDATE SET score = ?''',
                (assessment_id, item_number, score, score)
            )
        
        conn.commit()
        conn.close()
    
    def get_ratings(self, assessment_id: int) -> dict:
        """Get all ratings for an assessment as {item_number: score}."""
        rows = self._fetchall(
            'SELECT item_number, score FROM ratings WHERE assessment_id = ?',
            (assessment_id,)
        )
        return {r['item_number']: r['score'] for r in rows}
    
    # =========== OPEN RESPONSE OPERATIONS ===========
    
    def save_open_response(self, assessment_id: int, question_number: int, response_text: str):
        """Save or update an open response."""
        conn, cursor = self._execute(
            '''INSERT INTO open_responses (assessment_id, question_number, response_text)
               VALUES (?, ?, ?)
               ON CONFLICT (assessment_id, question_number) 
               DO UPDATE SET response_text = ?''',
            (assessment_id, question_number, response_text, response_text)
        )
        conn.commit()
        conn.close()
    
    def save_all_open_responses(self, assessment_id: int, responses: dict):
        """Save all open responses at once. responses = {question_number: response_text}"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        for question_number, response_text in responses.items():
            cursor.execute(
                '''INSERT INTO open_responses (assessment_id, question_number, response_text)
                   VALUES (?, ?, ?)
                   ON CONFLICT (assessment_id, question_number) 
                   DO UPDATE SET response_text = ?''',
                (assessment_id, question_number, response_text, response_text)
            )
        
        conn.commit()
        conn.close()
    
    def get_open_responses(self, assessment_id: int) -> dict:
        """Get all open responses for an assessment as {question_number: response_text}."""
        rows = self._fetchall(
            'SELECT question_number, response_text FROM open_responses WHERE assessment_id = ?',
            (assessment_id,)
        )
        return {r['question_number']: r['response_text'] for r in rows}
    
    # =========== REPORTING QUERIES ===========
    
    def get_participant_data(self, participant_id: int) -> dict:
        """Get complete data for a participant including PRE and POST scores."""
        participant = self.get_participant(participant_id)
        if not participant:
            return None
        
        cohort = self.get_cohort(participant['cohort_id'])
        assessments = self.get_assessments_for_participant(participant_id)
        
        data = {
            'participant': participant,
            'cohort': cohort,
            'pre': None,
            'post': None
        }
        
        if assessments['PRE'] and assessments['PRE'].get('completed_at'):
            data['pre'] = {
                'assessment': assessments['PRE'],
                'ratings': self.get_ratings(assessments['PRE']['id']),
                'open_responses': self.get_open_responses(assessments['PRE']['id'])
            }
        
        if assessments['POST'] and assessments['POST'].get('completed_at'):
            data['post'] = {
                'assessment': assessments['POST'],
                'ratings': self.get_ratings(assessments['POST']['id']),
                'open_responses': self.get_open_responses(assessments['POST']['id'])
            }
        
        return data
    
    def get_cohort_data(self, cohort_id: int) -> dict:
        """Get aggregated data for a cohort."""
        cohort = self.get_cohort(cohort_id)
        if not cohort:
            return None
        
        participants = self.get_participants_for_cohort(cohort_id)
        
        data = {
            'cohort': cohort,
            'participants': [],
            'pre_completed': 0,
            'post_completed': 0,
            'total': len(participants)
        }
        
        for p in participants:
            p_data = self.get_participant_data(p['id'])
            if p_data:
                data['participants'].append(p_data)
                if p_data['pre']:
                    data['pre_completed'] += 1
                if p_data['post']:
                    data['post_completed'] += 1
        
        return data
    
    def get_cohort_averages(self, cohort_id: int, assessment_type: str = 'POST') -> dict:
        """Get average scores per item for a cohort."""
        query = '''
            SELECT r.item_number, AVG(r.score) as avg_score, COUNT(*) as count
            FROM ratings r
            JOIN assessments a ON r.assessment_id = a.id
            JOIN participants p ON a.participant_id = p.id
            WHERE p.cohort_id = ? AND a.assessment_type = ? AND a.completed_at IS NOT NULL
            GROUP BY r.item_number
        '''
        rows = self._fetchall(query, (cohort_id, assessment_type))
        return {r['item_number']: {'avg': r['avg_score'], 'count': r['count']} for r in rows}
    
    def get_all_open_responses_for_cohort(self, cohort_id: int, assessment_type: str, question_number: int) -> list:
        """Get all open responses for a specific question across a cohort."""
        query = '''
            SELECT o.response_text, p.name
            FROM open_responses o
            JOIN assessments a ON o.assessment_id = a.id
            JOIN participants p ON a.participant_id = p.id
            WHERE p.cohort_id = ? AND a.assessment_type = ? AND o.question_number = ?
            AND a.completed_at IS NOT NULL AND o.response_text IS NOT NULL AND o.response_text != ''
        '''
        return self._fetchall(query, (cohort_id, assessment_type, question_number))
