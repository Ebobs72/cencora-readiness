"""
load_test_data.py — Loads a full synthetic test cohort into the Cencora database.

Matches the actual schema:
  cohorts (AUTOINCREMENT id)
  participants (cohort_id FK) 
  assessments (participant_id FK, assessment_type, access_token, completed_at)
  ratings (assessment_id FK, item_number, score)
  open_responses (assessment_id FK, question_number, response_text)
"""

import random
import secrets
from datetime import datetime, timedelta


def load_test_cohort(db):
    """Load a complete test cohort with PRE and POST data."""
    
    random.seed(42)
    
    PRE_DATE = datetime(2025, 11, 14)
    POST_DATE = datetime(2026, 2, 12)
    
    participants = [
        {"name": "Sarah Mitchell", "role": "Shift Manager", "email": "s.mitchell@cencora-test.com"},
        {"name": "James Okonkwo", "role": "Warehouse Operations Manager", "email": "j.okonkwo@cencora-test.com"},
        {"name": "Emma Zhao", "role": "Team Leader - Inbound", "email": "e.zhao@cencora-test.com"},
        {"name": "David Patel", "role": "Site Safety Lead", "email": "d.patel@cencora-test.com"},
        {"name": "Rachel Thompson", "role": "Distribution Centre Manager", "email": "r.thompson@cencora-test.com"},
        {"name": "Marcus Williams", "role": "Shift Manager", "email": "m.williams@cencora-test.com"},
        {"name": "Lisa Kowalski", "role": "Team Leader - Outbound", "email": "l.kowalski@cencora-test.com"},
        {"name": "Tom Hennessey", "role": "Transport & Logistics Manager", "email": "t.hennessey@cencora-test.com"},
        {"name": "Priya Sharma", "role": "Quality & Compliance Manager", "email": "p.sharma@cencora-test.com"},
        {"name": "Chris Barker", "role": "Team Leader - Pick & Pack", "email": "c.barker@cencora-test.com"},
        {"name": "Nina Osei", "role": "People & Training Coordinator", "email": "n.osei@cencora-test.com"},
        {"name": "Ryan Gallagher", "role": "Shift Manager", "email": "r.gallagher@cencora-test.com"},
    ]
    
    profiles = {
        "Sarah Mitchell": {
            "pre": {"Self": 4.5, "Practical": 3.5, "Professional": 3.8, "Team": 3.2, "Overall": 3.8},
            "growth": {"Self": 0.5, "Practical": 1.2, "Professional": 0.8, "Team": 1.0, "Overall": 0.8},
        },
        "James Okonkwo": {
            "pre": {"Self": 3.2, "Practical": 4.5, "Professional": 3.5, "Team": 4.0, "Overall": 3.8},
            "growth": {"Self": 1.0, "Practical": 0.5, "Professional": 1.0, "Team": 0.6, "Overall": 0.7},
        },
        "Emma Zhao": {
            "pre": {"Self": 2.8, "Practical": 3.0, "Professional": 2.5, "Team": 2.8, "Overall": 2.5},
            "growth": {"Self": 1.5, "Practical": 1.3, "Professional": 1.5, "Team": 1.2, "Overall": 1.8},
        },
        "David Patel": {
            "pre": {"Self": 3.0, "Practical": 3.5, "Professional": 3.2, "Team": 4.8, "Overall": 3.5},
            "growth": {"Self": 1.2, "Practical": 0.8, "Professional": 1.0, "Team": 0.3, "Overall": 0.8},
        },
        "Rachel Thompson": {
            "pre": {"Self": 4.2, "Practical": 4.0, "Professional": 4.5, "Team": 4.0, "Overall": 4.5},
            "growth": {"Self": 0.5, "Practical": 0.6, "Professional": 0.5, "Team": 0.8, "Overall": 0.5},
        },
        "Marcus Williams": {
            "pre": {"Self": 3.0, "Practical": 3.8, "Professional": 3.0, "Team": 3.5, "Overall": 3.2},
            "growth": {"Self": 1.3, "Practical": 0.7, "Professional": 1.2, "Team": 0.8, "Overall": 1.0},
        },
        "Lisa Kowalski": {
            "pre": {"Self": 3.5, "Practical": 3.2, "Professional": 3.0, "Team": 3.0, "Overall": 2.8},
            "growth": {"Self": 1.0, "Practical": 1.0, "Professional": 1.2, "Team": 1.0, "Overall": 1.5},
        },
        "Tom Hennessey": {
            "pre": {"Self": 4.0, "Practical": 4.2, "Professional": 3.0, "Team": 2.5, "Overall": 3.5},
            "growth": {"Self": 0.3, "Practical": 0.4, "Professional": 1.0, "Team": 1.5, "Overall": 0.8},
        },
        "Priya Sharma": {
            "pre": {"Self": 3.8, "Practical": 3.5, "Professional": 3.8, "Team": 3.5, "Overall": 3.2},
            "growth": {"Self": 0.7, "Practical": 0.8, "Professional": 0.7, "Team": 0.8, "Overall": 1.0},
        },
        "Chris Barker": {
            "pre": {"Self": 2.5, "Practical": 2.8, "Professional": 2.5, "Team": 3.0, "Overall": 2.2},
            "growth": {"Self": 1.2, "Practical": 1.5, "Professional": 1.3, "Team": 1.0, "Overall": 1.5},
        },
        "Nina Osei": {
            "pre": {"Self": 4.0, "Practical": 3.0, "Professional": 3.5, "Team": 4.2, "Overall": 3.8},
            "growth": {"Self": 0.5, "Practical": 1.2, "Professional": 0.8, "Team": 0.5, "Overall": 0.7},
        },
        "Ryan Gallagher": {
            "pre": {"Self": 4.8, "Practical": 4.5, "Professional": 4.2, "Team": 4.0, "Overall": 4.8},
            "growth": {"Self": -0.3, "Practical": 0.2, "Professional": 0.3, "Team": 0.5, "Overall": -0.2},
        },
    }
    
    # Item-to-indicator mapping
    item_indicators = {}
    for i in range(1, 7): item_indicators[i] = "Self"
    for i in range(7, 15): item_indicators[i] = "Practical"
    for i in range(15, 23): item_indicators[i] = "Professional"
    for i in range(23, 31): item_indicators[i] = "Team"
    for i in range(31, 33): item_indicators[i] = "Overall"
    
    item_focus = {}
    focus6 = ["Knowledge", "Knowledge", "Awareness", "Awareness", "Confidence", "Behaviour"]
    focus8 = ["Knowledge", "Knowledge", "Awareness", "Awareness", "Confidence", "Confidence", "Behaviour", "Behaviour"]
    for i in range(1, 7): item_focus[i] = focus6[i - 1]
    for i in range(7, 15): item_focus[i] = focus8[i - 7]
    for i in range(15, 23): item_focus[i] = focus8[i - 15]
    for i in range(23, 31): item_focus[i] = focus8[i - 23]
    item_focus[31] = "Confidence"
    item_focus[32] = "Confidence"
    
    # Open-ended responses
    pre_responses = {
        "Sarah Mitchell": [
            "Building something from the ground up - I've always inherited teams so this feels like a real opportunity to shape things from day one.",
            "Having difficult conversations. I tend to avoid conflict and I know that's going to hold me back if I don't address it.",
            "That we'll be under so much pressure to hit targets that the people stuff will get pushed aside. I've seen it happen before."
        ],
        "James Okonkwo": [
            "Getting the operations running smoothly. I love the challenge of building systems and processes from scratch.",
            "My communication skills - I'm great at the technical side but I know I need to get better at bringing people along with me.",
            "Honestly, whether the leadership team will give us enough time to bed things in before expecting full performance."
        ],
        "Emma Zhao": [
            "Leading my own team for the first time. I've been a strong individual contributor and I'm excited to step up.",
            "Everything! But specifically, I'd like to feel more confident in my ability to manage people rather than tasks.",
            "That people won't take me seriously because I'm younger than most of the team. I worry about credibility."
        ],
        "David Patel": [
            "Embedding safety culture from day one rather than retrofitting it. This is a rare opportunity to get it right from the start.",
            "The softer side of leadership - I know safety inside out but I want to get better at coaching and developing people.",
            "That safety will be seen as David's job rather than everyone's responsibility. Getting buy-in across all shifts worries me."
        ],
        "Rachel Thompson": [
            "Shaping the culture. I've managed large teams before but never had the chance to build from zero. This is what I've been working towards.",
            "Delegation - I know I hold on to too much. With a site this size I simply won't be able to do everything myself.",
            "The sheer scale of it. There are a lot of moving parts and I want to make sure nothing falls through the cracks."
        ],
        "Marcus Williams": [
            "Getting stuck in and making things happen. I'm a practical person and this kind of challenge energises me.",
            "Being more aware of how I come across. I've had feedback before that I can be too direct and not realise the impact.",
            "That the pace won't let up long enough for me to actually reflect on what I'm doing. I tend to just crack on."
        ],
        "Lisa Kowalski": [
            "Working in a purpose-built facility - the outbound processes will be much more efficient than what I'm used to.",
            "My confidence in speaking up in meetings with senior managers. I know my stuff but I clam up when challenged.",
            "Not being good enough. I was promoted quite quickly and sometimes I wonder if I'm really ready for this."
        ],
        "Tom Hennessey": [
            "Designing the transport and logistics operation from scratch. No legacy issues, no inherited problems.",
            "Working more collaboratively across departments. I tend to focus on my own area and I know I need to be more joined up.",
            "Getting pulled into other people's problems when I need to focus on getting my own operation right first."
        ],
        "Priya Sharma": [
            "Establishing quality standards before bad habits form. Prevention is always better than correction.",
            "Influencing people without formal authority. I need buy-in from operational managers who don't report to me.",
            "That compliance will be seen as a blocker rather than an enabler. I need to position quality as everyone's friend."
        ],
        "Chris Barker": [
            "To be honest, I'm still getting my head around the fact that I'm doing this. I was happy as a picker but my manager pushed me to apply.",
            "Basically everything about managing people. I know the warehouse inside out but leading a team is completely new to me.",
            "That I'll let people down. My team are relying on me and I don't really know what I'm doing yet."
        ],
        "Nina Osei": [
            "Supporting people through what will be a massive transition. I love helping people develop and grow.",
            "The process and systems side - I'm great with people but I need to get better at tracking, reporting and following through on data.",
            "That we'll be so busy with operations that training and development will get deprioritised. It always seems to be first on the chopping block."
        ],
        "Ryan Gallagher": [
            "I've done three site launches before so I know what to expect. Looking forward to using that experience here.",
            "I'm pretty confident across the board to be honest. Maybe just fine-tuning my approach for a new organisation's culture.",
            "Not really any concerns. I've been through this before and I know the playbook."
        ],
    }
    
    post_responses = {
        "Sarah Mitchell": [
            "The difficult conversations framework. Having a clear process to follow has made me much braver about tackling things head-on.",
            "I'll address issues early rather than letting them fester. The programme showed me that avoiding conflict IS the conflict.",
            "I still worry about targets vs people, but I feel much better equipped to hold both. The prioritisation tools really helped."
        ],
        "James Okonkwo": [
            "The DISC profiling was a revelation. Understanding that not everyone processes information the way I do has changed how I communicate.",
            "I'll tailor my communication based on who I'm talking to rather than assuming everyone wants the same level of detail I do.",
            "I'm more patient now. I realise that getting people right will ultimately get operations right - they're not competing priorities."
        ],
        "Emma Zhao": [
            "Realising that leadership isn't about having all the answers. The Values exercise gave me real clarity about what I stand for.",
            "I'll have regular one-to-ones with everyone - not just when there's a problem. And I'll delegate more instead of trying to prove myself by doing everything.",
            "The credibility concern hasn't gone away entirely but I feel so much more prepared. Knowing my values and having practical frameworks gives me something solid to stand on."
        ],
        "David Patel": [
            "The coaching conversation model. I've always told people what to do for safety - now I understand how to get them to think for themselves.",
            "More coaching, less telling. I'll ask questions before jumping in with solutions, especially in safety conversations.",
            "I'm more confident that I can influence safety culture through how I lead, not just through policies and procedures."
        ],
        "Rachel Thompson": [
            "The delegation framework was exactly what I needed. The Freedom Ladder has already changed how I brief my direct reports.",
            "Let go more. I've already started using the delegation model and the difference in my workload is noticeable. I need to trust my team more.",
            "The scale still concerns me but I feel like I have much better tools to manage it. The prioritisation matrix was immediately useful."
        ],
        "Marcus Williams": [
            "The feedback exercise was uncomfortable but powerful. Hearing how others experience my directness was eye-opening.",
            "I'll pause before responding, especially in high-pressure moments. The impact vs intent concept really stuck with me.",
            "I've started building in reflection time at the end of each shift. It's only 10 minutes but it's making a real difference to how I show up the next day."
        ],
        "Lisa Kowalski": [
            "The group exercises, honestly. Realising that other people have the same doubts made me feel less alone. And the confidence building was practical, not fluffy.",
            "I've committed to contributing at least once in every meeting I attend. Small step but it's already getting easier.",
            "I still have wobbles but the imposter syndrome has quieted down a lot. Knowing that feeling under-confident is normal for this stage really helped."
        ],
        "Tom Hennessey": [
            "The stakeholder mapping exercise. I've been so focused on my own department that I hadn't properly thought about all the dependencies.",
            "I'll set up regular cross-functional check-ins rather than waiting for problems to force collaboration.",
            "I'm less worried about getting pulled in because I now have better tools to manage my boundaries while still being collaborative."
        ],
        "Priya Sharma": [
            "The influencing styles work. Understanding that different people need different approaches to be persuaded was really practical.",
            "Position quality conversations as helping you succeed rather than checking up on you. Framing makes such a difference.",
            "I feel more confident about winning hearts and minds. The programme gave me practical tools rather than just theory."
        ],
        "Chris Barker": [
            "Everything, genuinely. But if I had to pick one thing, it's that being a good leader isn't about knowing everything - it's about creating the conditions for your team to do their best work.",
            "Actually have one-to-ones rather than just catching people on the floor. And use the feedback model rather than just hoping people know what I think.",
            "Night and day difference. I went in terrified and I'm coming out feeling like I can actually do this. I've got frameworks, I've got a buddy, and I know where to go for help."
        ],
        "Nina Osei": [
            "The structured approach to meetings and tracking. I now have a framework for ensuring development actions don't just get talked about - they get done.",
            "Implement a proper tracker for all people development actions, with follow-up dates. No more relying on memory and good intentions.",
            "I'm cautiously optimistic. The programme has given development real legitimacy with the operational managers. They can see it's not nice to have."
        ],
        "Ryan Gallagher": [
            "The Values exercise caught me off guard. I realised I've been running on autopilot from previous launches without really thinking about what matters here specifically.",
            "Listen more. I came in thinking I had all the answers and the DISC work showed me I've been steamrolling people without realising it.",
            "Turns out I had more blind spots than I thought. The programme was genuinely humbling - in a good way."
        ],
    }
    
    def gen_score(base, focus, is_post=False, growth=0):
        noise = random.gauss(0, 0.3)
        score = base + noise
        if focus == "Confidence":
            score -= 0.3
            if is_post: growth += 0.2
        elif focus == "Knowledge":
            score += 0.2
        if is_post:
            score += growth + random.gauss(0, 0.4)
        return max(1, min(6, round(score)))
    
    # ── Remove any previous test data ──
    # We tag test cohorts with a recognisable name
    
    conn = db.get_connection()
    cursor = conn.cursor()
    
    # Find and delete any existing test cohort
    cursor.execute("SELECT id FROM cohorts WHERE name = 'Test Cohort - Wave 1'")
    existing = cursor.fetchall()
    for row in existing:
        cohort_id_to_delete = row[0] if isinstance(row, (list, tuple)) else row['id']
        # Get participants
        cursor.execute("SELECT id FROM participants WHERE cohort_id = ?", (cohort_id_to_delete,))
        old_participants = cursor.fetchall()
        for op in old_participants:
            op_id = op[0] if isinstance(op, (list, tuple)) else op['id']
            # Get assessments
            cursor.execute("SELECT id FROM assessments WHERE participant_id = ?", (op_id,))
            old_assessments = cursor.fetchall()
            for oa in old_assessments:
                oa_id = oa[0] if isinstance(oa, (list, tuple)) else oa['id']
                cursor.execute("DELETE FROM ratings WHERE assessment_id = ?", (oa_id,))
                cursor.execute("DELETE FROM open_responses WHERE assessment_id = ?", (oa_id,))
            cursor.execute("DELETE FROM assessments WHERE participant_id = ?", (op_id,))
        cursor.execute("DELETE FROM participants WHERE cohort_id = ?", (cohort_id_to_delete,))
        cursor.execute("DELETE FROM cohorts WHERE id = ?", (cohort_id_to_delete,))
    conn.commit()
    
    # ── Create cohort (AUTOINCREMENT id) ──
    
    cursor.execute(
        "INSERT INTO cohorts (name, programme, description, start_date, end_date) VALUES (?, ?, ?, ?, ?)",
        ("Test Cohort - Wave 1", "Launch Readiness", "Synthetic test data for report testing",
         PRE_DATE.strftime("%Y-%m-%d"), POST_DATE.strftime("%Y-%m-%d"))
    )
    conn.commit()
    cohort_id = cursor.lastrowid
    
    # ── Create participants, assessments, ratings and responses ──
    
    ratings_count = 0
    responses_count = 0
    
    for p in participants:
        name = p["name"]
        profile = profiles[name]
        
        # Create participant (AUTOINCREMENT id)
        cursor.execute(
            "INSERT INTO participants (cohort_id, name, email, role) VALUES (?, ?, ?, ?)",
            (cohort_id, p["name"], p["email"], p["role"])
        )
        conn.commit()
        participant_id = cursor.lastrowid
        
        # Create PRE assessment
        pre_token = secrets.token_urlsafe(32)
        pre_ts = (PRE_DATE + timedelta(hours=random.randint(9, 17), minutes=random.randint(0, 59))).isoformat()
        cursor.execute(
            "INSERT INTO assessments (participant_id, assessment_type, access_token, started_at, completed_at) VALUES (?, 'PRE', ?, ?, ?)",
            (participant_id, pre_token, pre_ts, pre_ts)
        )
        conn.commit()
        pre_assessment_id = cursor.lastrowid
        
        # Create POST assessment
        post_token = secrets.token_urlsafe(32)
        post_ts = (POST_DATE + timedelta(hours=random.randint(9, 17), minutes=random.randint(0, 59))).isoformat()
        cursor.execute(
            "INSERT INTO assessments (participant_id, assessment_type, access_token, started_at, completed_at) VALUES (?, 'POST', ?, ?, ?)",
            (participant_id, post_token, post_ts, post_ts)
        )
        conn.commit()
        post_assessment_id = cursor.lastrowid
        
        # Generate and insert ratings for all 32 items
        for item_num in range(1, 33):
            indicator = item_indicators[item_num]
            focus = item_focus[item_num]
            base = profile["pre"][indicator]
            growth = profile["growth"][indicator]
            
            pre_score = gen_score(base, focus)
            post_score = gen_score(base, focus, is_post=True, growth=growth)
            if post_score - pre_score > 3:
                post_score = pre_score + 3
            
            cursor.execute(
                "INSERT INTO ratings (assessment_id, item_number, score) VALUES (?, ?, ?)",
                (pre_assessment_id, item_num, pre_score)
            )
            cursor.execute(
                "INSERT INTO ratings (assessment_id, item_number, score) VALUES (?, ?, ?)",
                (post_assessment_id, item_num, post_score)
            )
            ratings_count += 2
        
        # Insert open responses
        for q_num, response in enumerate(pre_responses[name], 1):
            cursor.execute(
                "INSERT INTO open_responses (assessment_id, question_number, response_text) VALUES (?, ?, ?)",
                (pre_assessment_id, q_num, response)
            )
            responses_count += 1
        
        for q_num, response in enumerate(post_responses[name], 1):
            cursor.execute(
                "INSERT INTO open_responses (assessment_id, question_number, response_text) VALUES (?, ?, ?)",
                (post_assessment_id, q_num, response)
            )
            responses_count += 1
    
    conn.commit()
    conn.close()
    
    return {
        "cohort": "Test Cohort - Wave 1",
        "participants": len(participants),
        "ratings": ratings_count,
        "open_responses": responses_count,
        "pre_date": PRE_DATE.strftime("%d %B %Y"),
        "post_date": POST_DATE.strftime("%d %B %Y"),
    }


def remove_test_cohort(db):
    """Remove the test cohort and all related data."""
    conn = db.get_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM cohorts WHERE name = 'Test Cohort - Wave 1'")
    cohorts = cursor.fetchall()
    
    for row in cohorts:
        cohort_id = row[0] if isinstance(row, (list, tuple)) else row['id']
        cursor.execute("SELECT id FROM participants WHERE cohort_id = ?", (cohort_id,))
        parts = cursor.fetchall()
        for p in parts:
            p_id = p[0] if isinstance(p, (list, tuple)) else p['id']
            cursor.execute("SELECT id FROM assessments WHERE participant_id = ?", (p_id,))
            assmnts = cursor.fetchall()
            for a in assmnts:
                a_id = a[0] if isinstance(a, (list, tuple)) else a['id']
                cursor.execute("DELETE FROM ratings WHERE assessment_id = ?", (a_id,))
                cursor.execute("DELETE FROM open_responses WHERE assessment_id = ?", (a_id,))
            cursor.execute("DELETE FROM assessments WHERE participant_id = ?", (p_id,))
        cursor.execute("DELETE FROM participants WHERE cohort_id = ?", (cohort_id,))
        cursor.execute("DELETE FROM cohorts WHERE id = ?", (cohort_id,))
    
    conn.commit()
    conn.close()
    return True
