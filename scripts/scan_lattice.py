import urllib.request
import json
import re

req_ids = [
    ('Member-of-Technical-Staff-I_R-103169', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103386', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103143', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103176', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103178', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103179', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103140', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103142-1', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103229', 'Chennai-India'),
    ('Member-of-Technical-Staff-I_R-103128', 'Chennai-India'),
    ('UEFI-BIOS-Engineer_R-103136', 'Chennai-India'),
    ('Silicon-Validation-Engineer_R-103155', 'Pune-India')
]

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)', 'Accept': 'application/json'}

results = []

for slug, loc in req_ids:
    url = f"https://latticesemi.wd5.myworkdayjobs.com/wday/cxs/latticesemi/latticesemiconductorscareers/job/{loc}/{slug}"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            d = json.loads(resp.read().decode('utf-8'))
            job_info = d.get('jobPostingInfo', {})
            title = job_info.get('title', '')
            desc = job_info.get('jobDescription', '')
            posted = job_info.get('startDate', '')
            clean = re.sub('<[^<]+?>', '\n', desc)
            lines = [l.strip() for l in clean.split('\n') if l.strip()]
            
            # Find the role overview
            role_text = ""
            for i, l in enumerate(lines):
                if any(x in l.upper() for x in ['THE ROLE', 'JOB SUMMARY', 'POSITION SUMMARY', 'WHAT YOU']):
                    role_text = "\n".join(lines[i:i+8])
                    break
            if not role_text and len(lines) > 5:
                role_text = "\n".join(lines[4:10])

            clean_url = f"https://latticesemi.wd5.myworkdayjobs.com/en-US/latticesemiconductorscareers/job/{loc}/{slug}"
            results.append({
                'title': title,
                'slug': slug,
                'loc': loc,
                'posted': posted,
                'url': clean_url,
                'role_text': role_text
            })
    except Exception as e:
        print(f"Error for {slug}: {e}")

print(f"\nSuccessfully parsed {len(results)} Lattice India roles:\n")
for r in results:
    print("=" * 70)
    print(f"TITLE: {r['title']}")
    print(f"LOCATION: {r['loc']}")
    print(f"DIRECT URL: {r['url']}")
    print("ROLE OVERVIEW:")
    print(r['role_text'])
    print("=" * 70)
