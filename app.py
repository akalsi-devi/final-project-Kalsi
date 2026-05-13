import os
import anthropic
import streamlit as st
from generate_pptx import generate_proposal_pptx

# ── API Clients ───────────────────────────────────────────────────────────────
client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
PEXELS_API_KEY = os.environ.get("PEXELS_API_KEY", "")
HEADSHOT_PATH = os.path.join(os.path.dirname(__file__), "prof_pic.png")

# ── Boilerplate ───────────────────────────────────────────────────────────────
BOILERPLATE = {
    "Property Management": """
PROPERTY MANAGEMENT SERVICES:
Cushman & Wakefield provides comprehensive property management services including:
- Day-to-day operations and tenant relations
- Lease administration and compliance monitoring
- Vendor management and contract oversight
- Monthly financial reporting and budget management
- Property inspections and maintenance coordination
- Emergency response coordination available 24/7
""",
    "Engineering Services": """
ENGINEERING SERVICES:
Cushman & Wakefield's engineering team delivers:
- Preventive and corrective maintenance programs
- HVAC, electrical, and plumbing systems management
- Energy management and sustainability initiatives
- Capital project planning and oversight
- Building automation system management
- Compliance with all applicable codes and regulations
""",
    "Janitorial Services": """
JANITORIAL SERVICES:
Cushman & Wakefield provides full-service janitorial and cleaning programs including:
- Daily, weekly, and monthly cleaning schedules
- Common area and tenant suite maintenance
- Green cleaning programs using environmentally responsible products
- Specialized floor care and carpet maintenance
- Window washing and exterior cleaning
- Emergency cleanup response
""",
    "Security Services": """
SECURITY SERVICES:
Cushman & Wakefield's security program includes:
- 24/7 security personnel and patrol services
- Access control system management
- CCTV monitoring and incident reporting
- Emergency response coordination
- Visitor management protocols
- Security risk assessments and recommendations
""",
    "Accounting Services": """
ACCOUNTING SERVICES:
Cushman & Wakefield provides full-service accounting and financial management including:
- Monthly and annual financial reporting
- Accounts payable and receivable management
- Budget preparation and variance analysis
- CAM reconciliations and tenant billing
- Audit support and compliance reporting
- Cash flow management and forecasting
""",
}

SYSTEM_PROMPT = """You are an expert commercial real estate proposal writer for Cushman & Wakefield Asset Services.
You write professional, client-ready first-draft proposals for property services.

Your proposals should:
- Be formal, professional, and client-appropriate in tone
- Be specific to the client name, property address, and selected services
- Follow a standard CRE proposal structure
- Never invent fees, prices, or specific terms not provided
- Flag any fields that need human review with [NEEDS REVIEW]
- Include a disclaimer at the end that this is an AI-generated draft

Always structure the proposal with these sections:
1. Cover / Introduction
2. Understanding of Your Needs
3. Scope of Services (one section per selected service)
4. Why Cushman & Wakefield
5. Next Steps
6. Disclaimer"""


def generate_proposal(client_name, property_address, property_type, services, additional_notes):
    boilerplate_context = "\n".join([BOILERPLATE[s] for s in services])
    user_message = f"""Please write a professional first-draft proposal with the following details:

CLIENT NAME: {client_name}
PROPERTY ADDRESS: {property_address}
PROPERTY TYPE: {property_type}
SELECTED SERVICES: {", ".join(services)}
ADDITIONAL NOTES: {additional_notes if additional_notes else "None"}

Use the following boilerplate content for the scope of services sections:
{boilerplate_context}

Write a complete, professional proposal draft."""

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}]
    )
    return message.content[0].text


# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="CRE Proposal Assistant", page_icon="🏢", layout="wide")

st.title("🏢 CRE Proposal Writing Assistant")
st.subheader("Cushman & Wakefield Asset Services")
st.caption("AI-powered first-draft generator. All outputs require human review before sending to clients.")

st.divider()

col1, col2 = st.columns(2)

with col1:
    client_name = st.text_input("Client Name *", placeholder="e.g. Acme Corporation")
    property_address = st.text_input("Property Address *", placeholder="e.g. 1234 Main St, Washington DC 20001")

with col2:
    property_type = st.selectbox("Property Type *", [
        "Office Building",
        "Industrial Warehouse",
        "Retail Strip Center",
        "Mixed-Use Property",
        "Medical Office",
        "Other"
    ])

services = st.multiselect(
    "Select Services *",
    list(BOILERPLATE.keys()),
    help="Select all services to include in the proposal"
)

additional_notes = st.text_area(
    "Additional Notes (optional)",
    placeholder="e.g. Renewal proposal for existing client, custom fee structure needed..."
)

st.divider()

if st.button("Generate Proposal", type="primary", use_container_width=True):
    if not client_name or not property_address or not services:
        st.error("Please fill in all required fields: Client Name, Property Address, and at least one Service.")
    else:
        with st.spinner("Generating your proposal draft..."):
            try:
                proposal = generate_proposal(
                    client_name, property_address, property_type, services, additional_notes
                )
                st.success("Proposal generated successfully!")
                st.divider()
                st.subheader("Generated Proposal Draft")
                st.warning("⚠️ This is an AI-generated draft. Review all details carefully before sending to a client.")
                st.markdown(proposal)

                st.download_button(
                    label="📄 Download as Text",
                    data=proposal,
                    file_name=f"proposal_{client_name.replace(' ', '_')}.txt",
                    mime="text/plain"
                )

            except Exception as e:
                st.error(f"Error generating proposal: {str(e)}")
                proposal = None

        if 'proposal' in locals() and proposal:
            with st.spinner("Building branded PowerPoint with property photo..."):
                try:
                    import tempfile
                    output_path = tempfile.mktemp(suffix=".pptx")

                    generate_proposal_pptx(
                        client_name=client_name,
                        property_address=property_address,
                        property_type=property_type,
                        services=services,
                        proposal_text=proposal,
                        pexels_api_key=PEXELS_API_KEY,
                        headshot_path=HEADSHOT_PATH if os.path.exists(HEADSHOT_PATH) else None,
                        output_path=output_path
                    )

                    with open(output_path, "rb") as f:
                        pptx_bytes = f.read()

                    st.download_button(
                        label="📊 Download Branded PowerPoint",
                        data=pptx_bytes,
                        file_name=f"CW_Proposal_{client_name.replace(' ', '_')}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    st.success("✅ PowerPoint ready with real property photo!")

                except Exception as e:
                    st.error(f"PowerPoint generation error: {str(e)}")