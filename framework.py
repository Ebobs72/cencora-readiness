#!/usr/bin/env python3
"""
Framework definition for the Cencora Launch Readiness Assessment.

Contains all indicators, items, focus tags, and open questions.
"""

# Rating scale
RATING_SCALE = {
    1: "Strongly Disagree",
    2: "Disagree", 
    3: "Slightly Disagree",
    4: "Slightly Agree",
    5: "Agree",
    6: "Strongly Agree"
}

# Indicators and their item ranges
INDICATORS = {
    'Self-Readiness': (1, 6),
    'Practical Readiness': (7, 14),
    'Professional Readiness': (15, 22),
    'Team Readiness': (23, 30),
}

# Overall readiness items (not part of main indicators)
OVERALL_ITEMS = (31, 32)

# Indicator descriptions for reports
INDICATOR_DESCRIPTIONS = {
    'Self-Readiness': "Personal awareness, values, presence and style",
    'Practical Readiness': "Time, delegation, listening, conversations and feedback",
    'Professional Readiness': "Communication, trust, meetings, goals and accountability",
    'Team Readiness': "Operational requirements, safety, change and resilience"
}

# Indicator colours (Cencora brand)
INDICATOR_COLOURS = {
    'Self-Readiness': '#461E96',      # Purple
    'Practical Readiness': '#00B4E6',  # Cyan
    'Professional Readiness': '#E6008C', # Magenta
    'Team Readiness': '#00DC8C'        # Green
}

# Focus tags - what each item measures
FOCUS_TAGS = {
    'Knowledge': "Understanding of concepts, processes and frameworks",
    'Awareness': "Recognition of own patterns, triggers and impact",
    'Confidence': "Self-belief and comfort in capability",
    'Behaviour': "Actions, habits and practices"
}

# All 32 items with text and focus tag
ITEMS = {
    # Self-Readiness (1-6)
    1: {
        'text': "I can clearly articulate my personal values and how they influence my decisions",
        'focus': 'Knowledge',
        'indicator': 'Self-Readiness'
    },
    2: {
        'text': "I understand my preferred working style and how it differs from others",
        'focus': 'Knowledge',
        'indicator': 'Self-Readiness'
    },
    3: {
        'text': "I recognise how my behaviour changes when I am under pressure",
        'focus': 'Awareness',
        'indicator': 'Self-Readiness'
    },
    4: {
        'text': "I project credibility and presence when communicating with others",
        'focus': 'Confidence',
        'indicator': 'Self-Readiness'
    },
    5: {
        'text': "I adapt my approach effectively when working with people who have different styles to me",
        'focus': 'Behaviour',
        'indicator': 'Self-Readiness'
    },
    6: {
        'text': "I actively seek feedback on my own performance and act on it",
        'focus': 'Behaviour',
        'indicator': 'Self-Readiness'
    },
    
    # Practical Readiness (7-14)
    7: {
        'text': "I prioritise my time effectively, focusing on high-value activities",
        'focus': 'Behaviour',
        'indicator': 'Practical Readiness'
    },
    8: {
        'text': "I protect time for important but non-urgent work rather than constantly firefighting",
        'focus': 'Behaviour',
        'indicator': 'Practical Readiness'
    },
    9: {
        'text': "I delegate tasks appropriately rather than taking on too much myself",
        'focus': 'Behaviour',
        'indicator': 'Practical Readiness'
    },
    10: {
        'text': "I understand how to match my delegation approach to the individual and the task",
        'focus': 'Knowledge',
        'indicator': 'Practical Readiness'
    },
    11: {
        'text': "I listen to fully understand before forming my response",
        'focus': 'Behaviour',
        'indicator': 'Practical Readiness'
    },
    12: {
        'text': "I address difficult issues directly rather than avoiding or delaying them",
        'focus': 'Behaviour',
        'indicator': 'Practical Readiness'
    },
    13: {
        'text': "I give feedback that is specific, constructive and focused on improvement",
        'focus': 'Behaviour',
        'indicator': 'Practical Readiness'
    },
    14: {
        'text': "I am comfortable receiving feedback, even when it is challenging to hear",
        'focus': 'Confidence',
        'indicator': 'Practical Readiness'
    },
    
    # Professional Readiness (15-22)
    15: {
        'text': "I communicate with clarity, adapting my message for different audiences",
        'focus': 'Behaviour',
        'indicator': 'Professional Readiness'
    },
    16: {
        'text': "I build trust quickly through consistency between my words and actions",
        'focus': 'Behaviour',
        'indicator': 'Professional Readiness'
    },
    17: {
        'text': "I understand what creates and what erodes trust in working relationships",
        'focus': 'Knowledge',
        'indicator': 'Professional Readiness'
    },
    18: {
        'text': "I run meetings that are focused, productive and worth people's time",
        'focus': 'Behaviour',
        'indicator': 'Professional Readiness'
    },
    19: {
        'text': "I conduct effective check-ins that go beyond just task updates",
        'focus': 'Behaviour',
        'indicator': 'Professional Readiness'
    },
    20: {
        'text': "I set clear goals so people understand what success looks like",
        'focus': 'Behaviour',
        'indicator': 'Professional Readiness'
    },
    21: {
        'text': "I take ownership of outcomes rather than attributing problems to external factors",
        'focus': 'Behaviour',
        'indicator': 'Professional Readiness'
    },
    22: {
        'text': "I hold myself and others accountable for commitments made",
        'focus': 'Behaviour',
        'indicator': 'Professional Readiness'
    },
    
    # Team Readiness (23-30)
    23: {
        'text': "I understand the key HR processes and policies relevant to my role",
        'focus': 'Knowledge',
        'indicator': 'Team Readiness'
    },
    24: {
        'text': "I feel equipped to handle common people management situations",
        'focus': 'Confidence',
        'indicator': 'Team Readiness'
    },
    25: {
        'text': "I model and actively promote safety-first behaviours",
        'focus': 'Behaviour',
        'indicator': 'Team Readiness'
    },
    26: {
        'text': "I speak up about safety concerns, even when it might be uncomfortable",
        'focus': 'Behaviour',
        'indicator': 'Team Readiness'
    },
    27: {
        'text': "I help my team understand and navigate change rather than just announcing it",
        'focus': 'Behaviour',
        'indicator': 'Team Readiness'
    },
    28: {
        'text': "I maintain my own effectiveness during periods of pressure and uncertainty",
        'focus': 'Behaviour',
        'indicator': 'Team Readiness'
    },
    29: {
        'text': "I recognise signs of stress in myself and take action before it escalates",
        'focus': 'Awareness',
        'indicator': 'Team Readiness'
    },
    30: {
        'text': "I support the wellbeing of my team, particularly during demanding periods",
        'focus': 'Behaviour',
        'indicator': 'Team Readiness'
    },
    
    # Overall Readiness (31-32)
    31: {
        'text': "Overall, I feel ready to perform effectively in my role",
        'focus': 'Confidence',
        'indicator': 'Overall'
    },
    32: {
        'text': "I am confident I can build a high-performing team from day one",
        'focus': 'Confidence',
        'indicator': 'Overall'
    },
}

# Open questions for PRE assessment
OPEN_QUESTIONS_PRE = {
    1: "What aspect of your new role are you most looking forward to?",
    2: "What is the one area where you would most like to build your confidence or skills?",
    3: "What concerns, if any, do you have about the launch period ahead?"
}

# Open questions for POST assessment
OPEN_QUESTIONS_POST = {
    1: "What was your most valuable takeaway from the programme?",
    2: "What will you do differently as a result of attending?",
    3: "Looking back at your pre-programme concerns, how do you feel now?"
}


def get_indicator_for_item(item_num: int) -> str:
    """Get the indicator name for a given item number."""
    for indicator, (start, end) in INDICATORS.items():
        if start <= item_num <= end:
            return indicator
    if OVERALL_ITEMS[0] <= item_num <= OVERALL_ITEMS[1]:
        return 'Overall'
    return None


def get_items_for_indicator(indicator: str) -> list:
    """Get all item numbers for a given indicator."""
    if indicator == 'Overall':
        return list(range(OVERALL_ITEMS[0], OVERALL_ITEMS[1] + 1))
    if indicator in INDICATORS:
        start, end = INDICATORS[indicator]
        return list(range(start, end + 1))
    return []


def get_items_by_focus(focus: str) -> list:
    """Get all item numbers with a given focus tag."""
    return [num for num, item in ITEMS.items() if item['focus'] == focus]


def get_focus_summary() -> dict:
    """Get count of items per focus tag."""
    summary = {focus: 0 for focus in FOCUS_TAGS}
    for item in ITEMS.values():
        summary[item['focus']] += 1
    return summary
