import re

def clean_body_text(body):
    """Remove Microsoft Teams help information and other unnecessary content"""
    if not body:
        return ""
    
    # Remove everything after "Need help?" line
    body = re.split(r"Need help\?.*?<https://aka\.ms/JoinTeamsMeeting\?omkt=.*?>", body, flags=re.DOTALL)[0]
    
    # Additional cleanup: remove common meeting footers
    patterns = [
        r"Microsoft Teams.*?(?:\r\n|\n).*?Join conversation",
        r"________________+.*$",
        r"Click here to join.*$",
        r"Join with a video conferencing.*$",
        r"Join Microsoft Teams Meeting.*$",
    ]
    
    for pattern in patterns:
        body = re.split(pattern, body, flags=re.DOTALL)[0]
    
    # Trim whitespace and remove extra blank lines
    body = re.sub(r'\n{3,}', '\n\n', body.strip())
    
    return body