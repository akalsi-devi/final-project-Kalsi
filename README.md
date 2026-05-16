AI-Powered First-Draft Generator for Asset Services
Live App: proposalmatic.streamlit.app
1. Context, User & Problem
Who the user is:
A senior administrator supporting Managing Directors at a commercial real estate firm (Cushman & Wakefield Asset Services, Mid-Atlantic region). This person is responsible for proposal writing, client communications, vendor coordination, and administrative operations across a portfolio of commercial properties in DC, Maryland, and Virginia.
What workflow this improves:
Every time the team pursues a new property management contract, a proposal must be drafted from scratch. This requires pulling together the right service descriptions, formatting the document consistently, customizing the language for the client, and getting it ready for MD review before sending. This process currently takes 30 to 90 minutes per proposal and is largely repetitive boilerplate work.
Why it matters:
Proposal turnaround time directly affects win rates in commercial real estate. Faster, more consistent proposals give the team a competitive edge and free the administrator to focus on higher-value work. A tool that reduces drafting time from 90 minutes to under 30 seconds represents a meaningful operational improvement at scale.

2. Solution & Design
What was built:
A Streamlit web app that takes client name, property address, property type, and selected service scopes as input, calls the Claude API to generate a structured first-draft proposal, and outputs both a text draft and a fully branded 7-slide PowerPoint presentation ready for client delivery. How it works:

The user fills out a short form in the Streamlit interface
The selected service scopes determine which boilerplate sections are included
The Claude API generates a professional proposal draft using a structured system prompt
A branded PowerPoint is automatically generated with a real property photo fetched from Pexels based on property type, populated client details, scope of services cards, and a bio slide
Both the text draft and PowerPoint are available for download
Key design choices:
No RAG: The project plan originally included RAG, but after instructor feedback the design was simplified. The boilerplate library is small enough (5 service types) to fit in a single context window, making RAG unnecessary and the system simpler and more reliable.
Structured system prompt: The system prompt contains all service boilerplate and explicit instructions not to invent fees, addresses, or terms not provided by the user. Fields requiring human input are flagged with [NEEDS REVIEW].
Human review built in: Every output includes a visible disclaimer. No proposal is sent to a client without MD review. The app is a drafting assistant, not an autonomous sending tool.
Branded PowerPoint output: The 7-slide deck follows Cushman & Wakefield branding (navy, red, white color scheme) and includes: cover slide, property showcase with real Pexels photo, proposal overview, scope of services, why C&W, next steps, and a bio slide.
Course concepts integrated:

Anatomy of an LLM call (Weeks 2-3): System prompt engineering with boilerplate context, output constraints, and structured formatting instructions. Prompt iterated 3 times to reduce hallucination and improve tone.
Evaluation design (Week 6): 10-case rubric scored on 4 dimensions with baseline comparison against manual Word template process.
3. Evaluation & Results
Baseline:
The current manual process: filling a Word template by hand with no AI assistance. Estimated time: 30 to 90 minutes per proposal. Output quality varies by drafter and is often inconsistent in formatting and tone.
Test set:
10 synthetic proposal requests covering different property types, service scopes, and client profiles:

Office building, full scope
Industrial warehouse, engineering only
Retail strip center, janitorial and engineering
Mixed-use property, full scope
Medical office, property management and accounting
Renewal proposal for existing client
New client, security and janitorial only
Edge case: missing property type (Other)
Edge case: single service selected
Edge case: all five services selected

What was found:
The AI draft matched or exceeded the manual baseline on all four dimensions. The most significant improvement is speed: proposals that previously took 30 to 90 minutes now generate in under 30 seconds. The structured system prompt with explicit anti-hallucination instructions eliminated invented content across all 10 test cases.
Where it still requires human judgment:

Custom fee structures or pricing not in the standard boilerplate
Renewal proposals where prior relationship context matters
Any jurisdiction-specific legal language
Final review and approval before client delivery

4. Artifact Snapshot
The Streamlit interface:
The user fills in client name, property address, property type, and selects which services to include. The selected services directly control which scope sections appear in the proposal output.
Sample input:

Client Name: Acme Corporation
Property Address: 1234 Main St, Washington DC 20001
Property Type: Office Building
Services: Property Management, Engineering Services, Accounting Services

Sample output (text draft excerpt):

Cushman & Wakefield Asset Services is pleased to submit this proposal to provide comprehensive property services for Acme Corporation at 1234 Main St, Washington DC 20001. With decades of experience managing premier commercial properties across the Mid-Atlantic region, we are confident in our ability to enhance the operational efficiency and long-term value of your asset...

PowerPoint output:
The branded 7-slide deck includes:

Cover slide with client name and date auto-populated
Property showcase slide with a real Pexels photo matched to the property type
Proposal overview with cleaned, formatted content
Scope of services cards for selected services only
Why Cushman & Wakefield
Next steps
Bio slide with headshot, title, and contact information


Setup & Usage Instructions
Prerequisites

Python 3.9 or higher
An Anthropic API key (get one at console.anthropic.com)
A Pexels API key (get one free at pexels.com/api)

Installation
1. Clone the repository:
bashgit clone https://github.com/akalsi-devi/final-project-Kalsi.git
cd final-project-Kalsi
2. Install dependencies:
bashpip3 install -r requirements.txt
3. Set your API keys:
bashexport ANTHROPIC_API_KEY="your-anthropic-key-here"
export PEXELS_API_KEY="your-pexels-key-here"
4. Run the app:
bashpython3 -m streamlit run app.py
5. Open in browser:
Navigate to http://localhost:8501
Usage

Fill in the client name and property address
Select the property type from the dropdown
Check the services to include in the proposal
Add any optional notes
Click Generate Proposal
Download the text draft or branded PowerPoint

Repository Structure
final-project-Kalsi/
├── app.py              # Main Streamlit application
├── generate_pptx.py    # Branded PowerPoint generator
├── prof_pic.png        # Bio slide headshot
├── requirements.txt    # Python dependencies
└── README.md           # This file
Dependencies
streamlit
anthropic
python-pptx
requests

This app is a drafting assistant only. All outputs require human review before sending to a client. No client data is stored; all inputs are session-only.


