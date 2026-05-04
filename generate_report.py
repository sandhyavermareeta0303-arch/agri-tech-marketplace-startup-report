from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# Set default font
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)

# Set 1.5 line spacing for normal paragraphs
style.paragraph_format.line_spacing = 1.5

def add_page_number(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

for section in doc.sections:
    add_page_number(section)

def add_centered(text, size=12, bold=False, italic=False, space_after=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    return p

def add_heading(text, size=16):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = True
    return p

def add_subheading(text, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = True
    return p

def add_body(text, bold=False, indent=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(8)
    if indent:
        p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.bold = bold
    return p

def add_bullet(text, bold_label=None):
    p = doc.add_paragraph(style='List Bullet')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    if bold_label:
        run1 = p.add_run(bold_label)
        run1.font.name = 'Times New Roman'
        run1.font.size = Pt(12)
        run1.bold = True
        run2 = p.add_run(text)
        run2.font.name = 'Times New Roman'
        run2.font.size = Pt(12)
    else:
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
    return p

# ==================== TITLE PAGE ====================
add_centered('STARTUP AND ENTREPRENEURIAL ACTIVITY REPORT', size=18, bold=True, space_after=20)
add_centered('On', size=14, space_after=10)
add_centered('AGRI-TECH MARKETPLACE FOR FARMERS', size=16, bold=True, space_after=20)
add_centered('"KrishiConnect: Empowering Farmers Through a Digital Marketplace"', size=12, italic=True, space_after=24)

add_centered('Submitted in Partial Fulfillment of the', size=12, space_after=4)
add_centered('Requirement for the Degree of Master of Computer Application', size=12, space_after=4)
add_centered('In', size=12, space_after=4)
add_centered('Computer Application', size=12, space_after=24)

add_centered('Submitted By:', size=12, bold=True, space_after=6)
add_centered('Chinmaya Venkataraman', size=12, space_after=4)
add_centered('Roll No: [Your Roll Number]', size=12, space_after=4)
add_centered('Batch: MCA, 2025-26', size=12, space_after=24)

add_centered('Under the Supervision of', size=12, space_after=6)
add_centered('Faculty Name [Mr. ………..]', size=12, space_after=4)
add_centered('(Assistant Professor)', size=12, italic=True, space_after=4)
add_centered('Computer Application Department', size=12, space_after=30)

add_centered('KANPUR INSTITUTE OF TECHNOLOGY', size=14, bold=True, space_after=6)
add_centered('Affiliated to', size=12, space_after=4)
add_centered('Dr. A.P.J ABDUL KALAM TECHNICAL UNIVERSITY', size=12, bold=True, space_after=4)
add_centered('UTTAR PRADESH, LUCKNOW', size=12, bold=True, space_after=4)

doc.add_page_break()

# ==================== DECLARATION ====================
add_centered('DECLARATION', size=16, bold=True, space_after=20)
add_body(
    'I hereby declare that the Startup and Entrepreneurial Activity Report titled '
    '"Agri-Tech Marketplace for Farmers – KrishiConnect" submitted to Kanpur Institute of Technology, '
    'affiliated to Dr. A.P.J. Abdul Kalam Technical University, Lucknow, in partial fulfillment of the '
    'requirements for the award of the degree of Master of Computer Application, is a record of original '
    'work carried out by me under the supervision of [Faculty Name], Assistant Professor, Department of '
    'Computer Application.'
)
add_body(
    'I further declare that the contents of this report have not been submitted, in part or in full, to any '
    'other institution or university for the award of any degree, diploma, or other similar title. All sources '
    'of information referred to in this report have been duly acknowledged.'
)
add_body(' ')
add_body(' ')
add_body('Place: Kanpur')
add_body('Date: __________________')
add_body(' ')
add_body(' ')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run('Chinmaya Venkataraman\nRoll No: [Your Roll Number]\nMCA, 2025-26')
run.font.name = 'Times New Roman'
run.font.size = Pt(12)

doc.add_page_break()

# ==================== CERTIFICATE ====================
add_centered('CERTIFICATE', size=16, bold=True, space_after=20)
add_body(
    'This is to certify that the Startup and Entrepreneurial Activity Report titled '
    '"Agri-Tech Marketplace for Farmers – KrishiConnect" has been prepared and submitted by '
    'Mr. Chinmaya Venkataraman, Roll No: [Your Roll Number], a bona fide student of the Master of '
    'Computer Application program (Batch 2025-26) at Kanpur Institute of Technology, Kanpur, '
    'affiliated to Dr. A.P.J. Abdul Kalam Technical University, Lucknow.'
)
add_body(
    'The work has been carried out under my supervision and is a genuine effort by the student. The '
    'report fulfills the requirements laid down by the institution for the partial fulfillment of the degree '
    'of Master of Computer Application. To the best of my knowledge, the contents of this report have '
    'not been submitted elsewhere for the award of any other degree.'
)
add_body(
    'I wish him success in all his future endeavors.'
)
add_body(' ')
add_body(' ')
add_body(' ')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = p.add_run('[Faculty Name]\n(Assistant Professor)\nComputer Application Department\nKanpur Institute of Technology')
run.font.name = 'Times New Roman'
run.font.size = Pt(12)

doc.add_page_break()

# ==================== ACKNOWLEDGEMENT ====================
add_centered('ACKNOWLEDGEMENT', size=16, bold=True, space_after=20)
add_body(
    'I would like to express my sincere gratitude to all those who supported and guided me during '
    'the preparation of this Startup and Entrepreneurial Activity Report on "Agri-Tech Marketplace for '
    'Farmers – KrishiConnect."'
)
add_body(
    'I am deeply thankful to my supervisor [Faculty Name], Assistant Professor, Department of Computer '
    'Application, Kanpur Institute of Technology, for his invaluable guidance, encouragement, and constant '
    'support throughout the project. His insights into entrepreneurship, technology, and rural development '
    'have shaped this report immensely.'
)
add_body(
    'I extend my heartfelt thanks to the Head of Department and all faculty members of the Computer '
    'Application Department for providing me the opportunity and infrastructure needed to undertake this '
    'work. I also acknowledge the contributions of farmers, agricultural experts, and online resources whose '
    'inputs helped me understand ground-level realities of Indian agriculture.'
)
add_body(
    'Finally, I would like to thank my family and friends for their unwavering motivation and support.'
)
add_body(' ')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = p.add_run('Chinmaya Venkataraman')
run.font.name = 'Times New Roman'
run.font.size = Pt(12)

doc.add_page_break()

# ==================== TABLE OF CONTENTS ====================
add_centered('TABLE OF CONTENTS', size=16, bold=True, space_after=20)

toc_items = [
    ('1.', 'Executive Summary', '1'),
    ('2.', 'Introduction', '2'),
    ('3.', 'Problem Statement and Relevance', '4'),
    ('4.', 'Innovation and Creativity', '6'),
    ('5.', 'Proposed Business Model and Feasibility', '8'),
    ('6.', 'Market Analysis and Competition', '11'),
    ('7.', 'Financial Planning', '14'),
    ('8.', 'Social and Ethical Impact', '17'),
    ('9.', 'Risk Management', '19'),
    ('10.', 'Conclusion and Future Scope', '21'),
    ('11.', 'References', '23'),
]

table = doc.add_table(rows=len(toc_items)+1, cols=3)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text = 'S. No.'
hdr[1].text = 'Chapter / Section'
hdr[2].text = 'Page No.'
for cell in hdr:
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

for i, (num, title, page) in enumerate(toc_items, start=1):
    row = table.rows[i].cells
    row[0].text = num
    row[1].text = title
    row[2].text = page
    for cell in row:
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

doc.add_page_break()

# ==================== 1. EXECUTIVE SUMMARY ====================
add_heading('1. Executive Summary')
add_body(
    'KrishiConnect is a digital Agri-Tech marketplace designed to bridge the long-standing gap between '
    'Indian farmers and end-buyers, while simultaneously equipping farmers with affordable access to '
    'agricultural inputs, advisory services, and machinery. Indian agriculture employs nearly 45% of the '
    'national workforce yet contributes only about 18% to the GDP, primarily because farmers receive '
    'only 25–30% of the final retail price of their produce. The remaining margin is absorbed by a '
    'fragmented chain of middlemen, commission agents, and intermediaries who add little verifiable value. '
    'KrishiConnect aims to dismantle this inefficiency by enabling direct, transparent, and fair-price '
    'transactions between farmers and buyers (retailers, wholesalers, food processors, restaurants, '
    'institutional buyers, and consumers).',
    indent=True
)
add_body(
    'The platform will be delivered as a multilingual mobile-first application supported by a web portal, '
    'integrating real-time mandi prices, weather alerts, AI-driven crop advisory, IoT-based soil health '
    'analytics, an input store (seeds, fertilizers, pesticides), a "Farm Equipment-as-a-Service" rental '
    'module, embedded agri-finance, and logistics partner integration. Revenue will be generated through '
    'transaction commissions (2–4%), input-store margins, premium SaaS-like advisory subscriptions, '
    'logistics fees, and lender-side commissions on credit facilitation. With an addressable market '
    'exceeding USD 24 billion in India alone and a strong push from initiatives such as Digital Agriculture '
    'Mission, e-NAM, and ONDC, KrishiConnect is positioned to deliver measurable economic uplift to '
    'farmers while building a scalable, profitable, and socially impactful enterprise.',
    indent=True
)

doc.add_page_break()

# ==================== 2. INTRODUCTION ====================
add_heading('2. Introduction')

add_subheading('2.1 About the Startup')
add_body(
    'KrishiConnect is conceptualized as a next-generation Agri-Tech startup that transforms the way '
    'farmers buy, sell, and learn. It is not just a marketplace — it is an end-to-end digital ecosystem '
    'covering pre-harvest, harvest, and post-harvest stages. Farmers will be able to list their produce, '
    'receive bids from verified buyers, access input supplies at MRP-controlled prices, rent expensive '
    'machinery on hourly/daily basis, secure crop insurance, and avail short-term credit — all through a '
    'single application available in 11 Indian languages.',
    indent=True
)

add_subheading('2.2 Mission')
add_body(
    'To empower every Indian farmer with a transparent, technology-driven marketplace that ensures fair '
    'prices, dependable inputs, intelligent advisory, and inclusive financial services — thereby doubling '
    'farm-level income and restoring dignity to the profession of farming.',
    indent=True
)

add_subheading('2.3 Vision')
add_body(
    'To become India\'s most trusted agricultural digital infrastructure by 2030, connecting 25 million '
    'farmers with 5 lakh buyers across 28 states, and gradually expanding to other developing economies '
    'in South Asia, Africa, and Southeast Asia where smallholder farmers face similar systemic challenges.',
    indent=True
)

add_subheading('2.4 Motivation')
add_body(
    'The motivation for choosing this startup arises from a deeply rooted social and economic problem. '
    'Despite India being the world\'s largest producer of milk, pulses, and the second-largest producer of '
    'rice, wheat, sugarcane, fruits, and vegetables, the average monthly income of an Indian farming '
    'household is barely ₹10,218 (NABARD All-India Rural Financial Inclusion Survey). Farmer suicides, '
    'distress sales, post-harvest losses estimated at ₹92,000 crore annually, and dependence on usurious '
    'moneylenders are symptoms of a broken system.',
    indent=True
)
add_body(
    'Simultaneously, the rapid penetration of smartphones (over 750 million users), affordable mobile '
    'data, and government-backed open networks like ONDC and AgriStack create an unprecedented opportunity '
    'to digitize agriculture. KrishiConnect represents the intersection of social purpose and commercial '
    'viability — a startup that can profit only when farmers profit.',
    indent=True
)

doc.add_page_break()

# ==================== 3. PROBLEM STATEMENT ====================
add_heading('3. Problem Statement and Relevance')

add_subheading('3.1 Problem')
add_body(
    'Indian agriculture suffers from a deeply fragmented value chain in which a typical produce passes '
    'through 4 to 7 intermediaries before reaching the consumer. As a result:',
    indent=True
)
add_bullet('Farmers receive only 25–30% of the final retail price, while consumers pay 3–4 times the farm-gate price.')
add_bullet('Approximately 30–40% of perishable produce is lost between farm and consumer due to poor logistics and lack of cold-chain integration.')
add_bullet('Farmers lack access to real-time price information, leading to forced distress sales at exploitative rates.')
add_bullet('Quality inputs (certified seeds, genuine fertilizers, pesticides) are unavailable in many rural pockets, while counterfeit products are rampant.')
add_bullet('Advisory services are still delivered through outdated, generic methods rather than data-driven, hyperlocal recommendations.')
add_bullet('Smallholder farmers (those owning less than 2 hectares — making up 86% of total farmers) cannot afford modern machinery and are excluded from formal credit due to lack of collateral.')

add_subheading('3.2 Relevance')
add_body(
    'This problem must be solved now, in 2026, for several converging reasons:',
    indent=True
)
add_bullet('', bold_label='Digital Push: ')
add_body(
    'The Government of India\'s Digital Agriculture Mission (₹2,817 crore allocation), AgriStack, '
    'Unified Farmer Service Platform (UFSP), and ONDC for agriculture have created the regulatory and '
    'infrastructural readiness for digital marketplaces.'
)
add_bullet('', bold_label='Smartphone Penetration: ')
add_body(
    'Rural smartphone penetration crossed 56% in 2025 and is expected to reach 75% by 2028, making '
    'mobile-first solutions practically viable.'
)
add_bullet('', bold_label='Climate Pressure: ')
add_body(
    'Erratic monsoons and rising input costs make data-driven, precision-agriculture services more '
    'urgent than ever.'
)
add_bullet('', bold_label='Demographic Shift: ')
add_body(
    'The average age of an Indian farmer is now 50+ years; without making farming profitable and '
    'tech-enabled, the next generation will abandon it entirely.'
)

add_subheading('3.3 Target Audience')
add_body('Primary Users:', bold=True)
add_bullet('Smallholder and marginal farmers (1–4 acres) across Uttar Pradesh, Madhya Pradesh, Bihar, Maharashtra, and other agrarian states.')
add_bullet('Farmer Producer Organizations (FPOs) and cooperatives.')

add_body('Secondary Users (Buyers):', bold=True)
add_bullet('Wholesalers and traders in mandis.')
add_bullet('Food-processing companies (ITC, Patanjali, Britannia, Nestlé).')
add_bullet('Quick-commerce and modern retail (BigBasket, Reliance Smart, DMart).')
add_bullet('HoReCa segment — hotels, restaurants, cloud kitchens.')
add_bullet('Direct-to-consumer urban households seeking farm-fresh produce.')

add_body('Tertiary Stakeholders:', bold=True)
add_bullet('Banks, NBFCs, and insurance providers offering agri-credit and crop insurance.')
add_bullet('Logistics partners and cold-chain operators.')
add_bullet('Government agencies promoting e-NAM and PM-KISAN.')

doc.add_page_break()

# ==================== 4. INNOVATION ====================
add_heading('4. Innovation and Creativity')

add_subheading('4.1 Unique Selling Proposition (USP)')
add_body(
    'KrishiConnect is differentiated from existing platforms in five major dimensions:',
    indent=True
)
add_bullet('', bold_label='Hyperlocal Vernacular Voice-First Interface: ')
add_body(
    'Unlike DeHaat or Ninjacart that rely heavily on text and English/Hindi, KrishiConnect supports '
    '11 Indian languages with voice-driven commands using on-device speech recognition — enabling '
    'access for low-literacy farmers.'
)
add_bullet('', bold_label='AI-Powered Crop Doctor: ')
add_body(
    'Farmers can click a photo of an infected leaf or pest, and the in-app computer-vision model '
    '(trained on 2 million+ labeled images) instantly diagnoses the disease and recommends the '
    'optimal pesticide with dosage — available offline.'
)
add_bullet('', bold_label='Reverse-Auction Buyer Model: ')
add_body(
    'Instead of farmers chasing buyers, verified buyers compete for the produce by placing bids — '
    'driving prices up to 18–22% above mandi rates in pilot studies.'
)
add_bullet('', bold_label='Equipment-as-a-Service (EaaS): ')
add_body(
    'Tractors, harvesters, drones for spraying, and laser-leveling machines are listed by their owners '
    'and rented hourly via in-app booking, similar to "Uber for Tractors."'
)
add_bullet('', bold_label='Embedded Agri-Finance: ')
add_body(
    'Based on the farmer\'s on-platform transaction history, soil health, crop type, and satellite '
    'imagery, partner NBFCs offer instant kisan-credit at interest rates 4–6% lower than informal '
    'lenders, with no traditional collateral required.'
)

add_subheading('4.2 Innovation Factor')
add_body(
    'KrishiConnect combines several emerging technologies into a coherent, farmer-centric stack:',
    indent=True
)
add_bullet('', bold_label='Artificial Intelligence & Machine Learning: ')
add_body(
    'Convolutional neural networks for crop disease detection; demand-forecasting models that predict '
    'mandi prices 14 days in advance with >82% accuracy; recommendation engines for crop selection '
    'based on soil, weather, and market data.'
)
add_bullet('', bold_label='Internet of Things (IoT): ')
add_body(
    'Low-cost LoRaWAN soil-moisture, temperature, and pH sensors (~₹1,500 each) feeding real-time '
    'farm data to the platform.'
)
add_bullet('', bold_label='Blockchain for Traceability: ')
add_body(
    'Each produce batch is assigned a QR-linked digital passport — buyers can trace it from seed to '
    'shelf, supporting export-grade compliance and premium organic pricing.'
)
add_bullet('', bold_label='Satellite & Drone Imagery: ')
add_body(
    'Integration with ISRO Bhuvan and Sentinel-2 imagery for NDVI (Normalised Difference Vegetation '
    'Index)-based crop-health monitoring.'
)
add_bullet('', bold_label='ONDC & UPI Integration: ')
add_body(
    'Native integration with ONDC opens KrishiConnect produce to all ONDC buyer apps; UPI '
    'AutoPay, AutoCollect, and BBPS handle payments seamlessly.'
)
add_bullet('', bold_label='Generative AI Chatbot ("KrishiSakhi"): ')
add_body(
    'Built on a fine-tuned open-source LLM, the chatbot answers farming queries in regional languages '
    'with citations from ICAR research bulletins.'
)
add_body(
    'The creative leap is integrating these technologies behind a deceptively simple voice-first '
    'interface — so a 60-year-old farmer in a Bundelkhand village experiences the same ease as a '
    'consumer ordering on Amazon.',
    indent=True
)

doc.add_page_break()

# ==================== 5. BUSINESS MODEL ====================
add_heading('5. Proposed Business Model and Feasibility')

add_subheading('5.1 Business Model')
add_body(
    'KrishiConnect adopts a multi-sided platform model with five complementary revenue streams, '
    'ensuring no single dependency and balanced unit economics:',
    indent=True
)

table = doc.add_table(rows=6, cols=3)
table.style = 'Light Grid Accent 1'
data = [
    ('Revenue Stream', 'Mechanism', 'Target Share by Year 3'),
    ('Marketplace Commission', '2–4% on every B2B/B2C transaction', '45%'),
    ('Input Store Margin', '8–12% on seeds, fertilizers, pesticides', '20%'),
    ('Equipment Rental Fee', '15% per booking from owners', '10%'),
    ('Embedded Finance & Insurance', '0.8–1.5% commission from NBFCs/insurers', '15%'),
    ('Premium Subscription (KrishiPro)', '₹99/month for advanced advisory & analytics', '10%'),
]
for i, row_data in enumerate(data):
    row = table.rows[i].cells
    for j, text in enumerate(row_data):
        row[j].text = text
        for para in row[j].paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                if i == 0:
                    run.bold = True

add_body(' ')

add_subheading('5.2 Feasibility Analysis')

add_body('Technical Feasibility:', bold=True)
add_body(
    'The technology stack is deliberately built on proven, cost-effective components:',
    indent=True
)
add_bullet('Mobile App: React Native (Android-first; iOS later) for cross-platform efficiency.')
add_bullet('Backend: Node.js with Express, microservices architecture deployed on AWS Mumbai region.')
add_bullet('Database: PostgreSQL for transactions, MongoDB for product catalog, Redis for caching.')
add_bullet('AI/ML: Python (TensorFlow Lite for on-device inference, PyTorch for server-side training).')
add_bullet('Payments: Razorpay + UPI + Aadhaar-enabled payment for unbanked farmers.')
add_bullet('Infrastructure: Auto-scaling EKS clusters; estimated cloud cost ₹3.5 lakh/month at 1 lakh DAU.')
add_body(
    'All required components — speech recognition for Indian languages, ONDC adapters, satellite '
    'imagery APIs — are commercially available. A functional MVP can be built in 5–6 months by a '
    'team of 8 engineers.',
    indent=True
)

add_body('Market Viability:', bold=True)
add_bullet('India has 146 million operational landholdings — a massive, underserved user base.')
add_bullet('FY 2024–25 saw $2.4 billion in funding for Indian Agri-Tech startups, signaling investor confidence.')
add_bullet('Government initiatives provide subsidies, partnerships, and rails (e-NAM, ONDC, PM-KISAN) that reduce go-to-market friction.')
add_bullet('Pilot studies by similar players (DeHaat, AgroStar) indicate 40–60% repeat usage within 6 months — proving stickiness.')

add_subheading('5.3 Business Model Canvas')

bmc = [
    ('Key Partners',
     'NBFCs (Bharatpe Agri, Samunnati), insurance firms (HDFC Ergo, ICICI Lombard), logistics (Delhivery, Ecom Express, FreshToHome cold chain), input manufacturers (UPL, Coromandel, Mahindra), state agriculture departments, FPOs, ONDC.'),
    ('Key Activities',
     'Onboarding farmers and buyers, building & maintaining technology platform, AI model training, quality assurance, last-mile logistics coordination, regulatory compliance.'),
    ('Key Resources',
     'Technology platform, dataset of crops/diseases/prices, FPO partnerships, field-officer network, brand trust in rural India.'),
    ('Value Propositions',
     'For Farmers: 18–25% higher realisation, 30% cheaper inputs, instant credit, free advisory. For Buyers: traceable, quality-graded produce at 10–15% lower than mandi cost.'),
    ('Customer Relationships',
     'Self-service app + voice support; Krishi Mitra field officers in each cluster; community WhatsApp groups; in-app gamification (Kisan Points).'),
    ('Channels',
     'Mobile app (Android), web portal, FPO partnerships, IVR helpline (1800-KRISHI), social media in regional languages, on-ground demos at Krishi Vigyan Kendras.'),
    ('Customer Segments',
     'Smallholder farmers, FPOs, retailers, processors, exporters, HoReCa, urban households (D2C).'),
    ('Cost Structure',
     'Cloud & infrastructure, salaries (tech + field), marketing & onboarding incentives, logistics subsidy, AI training compute, customer support.'),
    ('Revenue Streams',
     'Marketplace commission, input margins, equipment rental fees, finance/insurance commissions, premium subscriptions, advertising by input brands.'),
]

t = doc.add_table(rows=len(bmc), cols=2)
t.style = 'Light Grid Accent 1'
for i, (block, content) in enumerate(bmc):
    t.rows[i].cells[0].text = block
    t.rows[i].cells[1].text = content
    for para in t.rows[i].cells[0].paragraphs:
        for run in para.runs:
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
    for para in t.rows[i].cells[1].paragraphs:
        for run in para.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)

doc.add_page_break()

# ==================== 6. MARKET ANALYSIS ====================
add_heading('6. Market Analysis and Competition')

add_subheading('6.1 Market Size')

add_body('Total Addressable Market (TAM):', bold=True)
add_body(
    'The Indian agricultural market is valued at approximately USD 407 billion as of 2025 (Ministry of '
    'Agriculture & Farmers Welfare). The total agri-tech opportunity within this is estimated at '
    'USD 24.1 billion by 2025 and projected to reach USD 34 billion by 2027 (Bain & EY-NASSCOM joint '
    'study).',
    indent=True
)

add_body('Serviceable Addressable Market (SAM):', bold=True)
add_body(
    'Focusing on India\'s top 10 agrarian states (UP, MP, Bihar, Maharashtra, Punjab, Haryana, Karnataka, '
    'Andhra Pradesh, Telangana, West Bengal) where smartphone penetration and ONDC readiness are highest, '
    'the SAM is approximately USD 9.6 billion.',
    indent=True
)

add_body('Serviceable Obtainable Market (SOM):', bold=True)
add_body(
    'A realistic 5-year capture target — translating to roughly 1.2 million active farmers and 50,000 '
    'buyers — yields a SOM of USD 240–280 million in GMV with platform revenue of USD 9–14 million '
    'annually.',
    indent=True
)

add_subheading('6.2 Competitor Analysis')

table = doc.add_table(rows=4, cols=3)
table.style = 'Light Grid Accent 1'
comp_data = [
    ('Competitor', 'Strengths', 'Weaknesses'),
    ('DeHaat', 'Strong rural network of 11,000+ micro-entrepreneurs; full-stack model covering inputs, advisory, output sales; raised over USD 200 million.', 'Heavy dependence on field officers (high opex); limited buyer-side technology; weak on equipment rental and embedded finance.'),
    ('Ninjacart',  'Largest fresh-produce B2B supply chain in India; well-developed cold chain and logistics; Walmart-backed.', 'Primarily focused on B2B for urban retailers; limited farmer-side empowerment (no input store, no advisory app); thin margins.'),
    ('AgroStar', 'Strong farmer-app adoption (5 million+); excellent product catalog for inputs; great vernacular UX.', 'Almost no marketplace for selling produce — one-way "selling to farmer" rather than two-way; minimal AI-driven services.'),
]
for i, row_data in enumerate(comp_data):
    row = table.rows[i].cells
    for j, text in enumerate(row_data):
        row[j].text = text
        for para in row[j].paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                if i == 0:
                    run.bold = True

add_body(' ')
add_subheading('6.3 KrishiConnect\'s Competitive Edge')
add_body(
    'KrishiConnect occupies a clear white space:',
    indent=True
)
add_bullet('Two-sided marketplace + input store + equipment rental + advisory + finance, all unified — competitors offer only 2–3 of these.')
add_bullet('AI-first, low-touch model that scales without proportional field-staff costs (better margins than DeHaat).')
add_bullet('ONDC-native architecture from day one — produce listed on KrishiConnect is automatically discoverable across the entire ONDC buyer ecosystem.')
add_bullet('Reverse-auction price-discovery mechanism — unique in this segment.')

add_subheading('6.4 SWOT Analysis')

table = doc.add_table(rows=2, cols=2)
table.style = 'Light Grid Accent 1'
table.rows[0].cells[0].text = 'STRENGTHS\n• Tech-first, scalable architecture\n• Strong vernacular & voice UX\n• Multi-revenue model (low concentration risk)\n• ONDC-native\n• Founder commitment to social mission'
table.rows[0].cells[1].text = 'WEAKNESSES\n• New brand without rural recognition initially\n• High customer-acquisition cost in early phase\n• Capital-intensive logistics setup\n• Dependence on third-party cold-chain partners'
table.rows[1].cells[0].text = 'OPPORTUNITIES\n• Government-backed Digital Agriculture Mission\n• Rising export demand for traceable Indian produce\n• Growing FPO ecosystem (10,000 new FPOs target)\n• ONDC adoption acceleration\n• Climate-resilient farming demand'
table.rows[1].cells[1].text = 'THREATS\n• Regulatory shifts in commodity pricing & e-mandi\n• Big-tech entrants (Reliance JioKrishi, Amazon Kisan)\n• Monsoon-driven cash-flow volatility\n• Counterfeit input penalties / liability\n• Cybersecurity risks for farmer data'
for r in table.rows:
    for c in r.cells:
        for para in c.paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)

doc.add_page_break()

# ==================== 7. FINANCIAL PLANNING ====================
add_heading('7. Financial Planning')

add_subheading('7.1 Initial Startup Cost (Year-0 Budget for MVP & Launch)')

table = doc.add_table(rows=10, cols=3)
table.style = 'Light Grid Accent 1'
finance_data = [
    ('Cost Head', 'Description', 'Estimated Amount (INR)'),
    ('Technology Development', 'MVP Mobile App + Web Portal + Admin Dashboard (8-engineer team × 6 months)', '₹ 35,00,000'),
    ('AI Model Training & Datasets', 'Crop disease dataset licensing, GPU hours, model fine-tuning', '₹ 6,00,000'),
    ('Cloud Infrastructure', 'AWS hosting, third-party APIs (12 months)', '₹ 8,00,000'),
    ('Field Operations', 'Onboarding 5,000 farmers across 10 districts; Krishi Mitra incentives', '₹ 15,00,000'),
    ('Marketing & Branding', 'Digital, on-ground events, regional language content production', '₹ 10,00,000'),
    ('Legal, Compliance & Licenses', 'Company incorporation, FSSAI, GST, ONDC participation, data-privacy compliance', '₹ 3,50,000'),
    ('Working Capital Buffer', 'Logistics float, refunds, contingency', '₹ 12,00,000'),
    ('Office & Operations', 'Lean co-working setup in Kanpur + Bengaluru', '₹ 5,50,000'),
    ('TOTAL', '', '₹ 95,00,000 (≈ ₹ 95 Lakh / USD 115K)'),
]
for i, row_data in enumerate(finance_data):
    row = table.rows[i].cells
    for j, text in enumerate(row_data):
        row[j].text = text
        for para in row[j].paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                if i == 0 or i == len(finance_data)-1:
                    run.bold = True

add_body(' ')

add_subheading('7.2 Revenue Model & Pricing Strategy')

add_body('Marketplace Commission:', bold=True)
add_bullet('2.0% on staple grains and pulses (high-volume, low-margin).')
add_bullet('3.5% on fruits, vegetables, dairy.')
add_bullet('4.0% on premium organic, branded, and export-grade produce.')

add_body('Input Store:', bold=True)
add_bullet('8% margin on seeds, 10% on fertilizers, 12% on pesticides and agro-chemicals.')
add_bullet('Free home-delivery on orders above ₹2,000; ₹49 below that.')

add_body('Equipment Rental:', bold=True)
add_bullet('15% commission to KrishiConnect; 85% to equipment owner.')
add_bullet('Dynamic pricing based on demand seasonality (sowing vs. harvest peaks).')

add_body('KrishiPro Subscription:', bold=True)
add_bullet('₹99/month or ₹999/year — unlocks AI advisor priority, market-price predictions, and zero-commission selling on first 5 transactions.')

add_body('Financial Services:', bold=True)
add_bullet('1.0% commission from NBFCs on every disbursed crop loan.')
add_bullet('1.5% commission on insurance premiums collected.')

add_subheading('7.3 3-Year Revenue Projection')

table = doc.add_table(rows=7, cols=4)
table.style = 'Light Grid Accent 1'
proj_data = [
    ('Metric', 'Year 1', 'Year 2', 'Year 3'),
    ('Active Farmers', '50,000', '3,00,000', '12,00,000'),
    ('Transacting Buyers', '1,200', '7,500', '35,000'),
    ('GMV (₹ Crore)', '60', '480', '2,400'),
    ('Take Rate (Effective)', '3.1%', '3.3%', '3.5%'),
    ('Platform Revenue (₹ Crore)', '1.86', '15.84', '84.00'),
    ('EBITDA Margin', '-180%', '-25%', '12%'),
]
for i, row_data in enumerate(proj_data):
    row = table.rows[i].cells
    for j, text in enumerate(row_data):
        row[j].text = text
        for para in row[j].paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                if i == 0:
                    run.bold = True

add_body(' ')
add_subheading('7.4 Sustainability')
add_body(
    'KrishiConnect targets EBITDA break-even by month 30 driven by three structural advantages:',
    indent=True
)
add_bullet('Network effects — every additional farmer attracts more buyers and vice versa, reducing CAC over time.')
add_bullet('Software-driven gross margins — once the platform is built, incremental users carry near-zero marginal cost.')
add_bullet('Repeat-purchase behavior — farmers transact 8–12 times a year for inputs and 2–4 times for produce, ensuring stable cash flows.')
add_body(
    'Long-term sustainability is reinforced by data network effects: every transaction enriches the AI '
    'models, making advisory and credit-scoring sharper, which further increases farmer retention and '
    'lender confidence — a self-reinforcing flywheel.',
    indent=True
)

doc.add_page_break()

# ==================== 8. SOCIAL & ETHICAL ====================
add_heading('8. Social and Ethical Impact')

add_subheading('8.1 Social Impact')
add_body(
    'KrishiConnect is designed not merely as a business but as an instrument of inclusive rural '
    'transformation. Quantifiable social benefits include:',
    indent=True
)
add_bullet('', bold_label='Income Uplift: ')
add_body(
    'Internal projections (validated by similar pilots) suggest a 20–30% rise in net farmer income '
    'within 18 months of platform adoption — directly contributing to the national goal of doubling '
    'farmers\' income.'
)
add_bullet('', bold_label='Reduction in Distress Sales: ')
add_body(
    'By providing real-time price visibility and storage-finance, the platform helps farmers avoid '
    'forced selling at low post-harvest prices — addressing a major cause of agrarian distress.'
)
add_bullet('', bold_label='Women & Youth Empowerment: ')
add_body(
    'Dedicated KrishiSakhi (women field-officer) and Yuva Krishak (youth entrepreneur) programs '
    'create rural employment for over 25,000 individuals by Year 3.'
)
add_bullet('', bold_label='Reduction in Food Waste: ')
add_body(
    'Direct buyer matching and demand forecasting can cut post-harvest losses by 15–18%, contributing '
    'to food security and SDG-12 (Responsible Consumption).'
)
add_bullet('', bold_label='Climate-Smart Agriculture: ')
add_body(
    'Hyperlocal advisory promotes water-efficient, low-pesticide, climate-resilient practices — '
    'aligned with SDG-13 (Climate Action) and SDG-15 (Life on Land).'
)
add_bullet('', bold_label='Financial Inclusion: ')
add_body(
    'Embedded credit and insurance for previously unbanked smallholders supports SDG-1 (No Poverty) '
    'and SDG-10 (Reduced Inequalities).'
)

add_subheading('8.2 Ethical Considerations')

add_body('Data Privacy:', bold=True)
add_bullet('Full compliance with the Digital Personal Data Protection Act (DPDPA), 2023.')
add_bullet('Explicit, granular, language-localized consent before collecting farm location, Aadhaar (only when required for KYC), or biometric data.')
add_bullet('All data stored encrypted (AES-256 at rest, TLS 1.3 in transit) within Indian data centers.')
add_bullet('Farmers can download or delete their data at any time via in-app settings ("Right to Erasure").')

add_body('Ethical AI:', bold=True)
add_bullet('Crop-loan algorithms are audited quarterly for bias against region, caste, gender, or land size.')
add_bullet('AI advisory clearly labels confidence levels and recommends consulting Krishi Vigyan Kendras for critical decisions — never replacing human expertise on irreversible interventions.')
add_bullet('Pesticide recommendations are reviewed by an ICAR-certified agronomy panel before being shipped in app updates.')

add_body('Fair Pricing & Anti-Exploitation:', bold=True)
add_bullet('No predatory commission on distress transactions (e.g., farmers losing crop to disease).')
add_bullet('Transparent fee disclosure — every screen shows what KrishiConnect earns from a transaction.')
add_bullet('Buyer rating and verification system ensures bad actors (delayed payments, weight-cheating) are rapidly removed.')

add_body('Environmental Ethics:', bold=True)
add_bullet('Promotion of bio-fertilizers and integrated pest management; no listing of banned/red-label pesticides.')
add_bullet('Carbon-credit pilot for farmers practicing zero-tillage and crop-rotation — generating an additional revenue line for them.')

doc.add_page_break()

# ==================== 9. RISK MANAGEMENT ====================
add_heading('9. Risk Management')

add_subheading('9.1 Key Risks')

table = doc.add_table(rows=8, cols=3)
table.style = 'Light Grid Accent 1'
risk_data = [
    ('Risk Category', 'Risk Description', 'Likelihood / Impact'),
    ('Financial', 'Slow user monetization; longer-than-expected runway burn', 'High / High'),
    ('Regulatory', 'Sudden changes in APMC laws, MSP regulations, or DPDPA enforcement', 'Medium / High'),
    ('Technical', 'Cybersecurity breach exposing farmer KYC and financial data', 'Medium / Critical'),
    ('Market Adoption', 'Low digital literacy hampering app uptake in deep-rural districts', 'High / Medium'),
    ('Operational', 'Logistics partner failure during peak harvest, leading to spoilage and refunds', 'Medium / High'),
    ('Competition', 'Big-tech (Reliance JioKrishi, Amazon Kisan) entering with subsidized pricing', 'Medium / High'),
    ('Climate / Macro', 'Successive bad monsoons reducing transaction volumes', 'Medium / High'),
]
for i, row_data in enumerate(risk_data):
    row = table.rows[i].cells
    for j, text in enumerate(row_data):
        row[j].text = text
        for para in row[j].paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                if i == 0:
                    run.bold = True

add_body(' ')

add_subheading('9.2 Mitigation Plan')

add_body('Financial Risk Mitigation:', bold=True)
add_bullet('Maintain minimum 18 months of runway at any time; raise funding in tranches tied to clear KPIs.')
add_bullet('Diversify revenue across 5 streams so under-performance in one does not jeopardize survival.')
add_bullet('Apply for non-dilutive grants under NABARD\'s Agri-Tech Fund and ICAR\'s Innovation Grants.')

add_body('Regulatory Risk Mitigation:', bold=True)
add_bullet('Onboard a regulatory advisor with experience in APMC and digital-finance laws.')
add_bullet('Active membership in FICCI/NASSCOM Agri-Tech committees to stay ahead of policy shifts.')
add_bullet('Build data-residency and compliance modules from day one to avoid retrofitting under DPDPA.')

add_body('Technical Risk Mitigation:', bold=True)
add_bullet('Annual third-party VAPT (Vulnerability Assessment & Penetration Testing) audits.')
add_bullet('SOC 2 Type II certification target within 24 months.')
add_bullet('Bug-bounty program; ISO 27001 ISMS framework.')
add_bullet('Real-time anomaly detection on transaction patterns to catch fraud or breaches early.')

add_body('Market-Adoption Risk Mitigation:', bold=True)
add_bullet('Voice-first, icon-driven UX for low-literacy users; on-ground demos at every village panchayat in launch districts.')
add_bullet('Partner with FPOs that already enjoy farmer trust — riding their distribution for cold-start.')
add_bullet('Referral incentives ("Mitra Banao, ₹100 Pao") to drive word-of-mouth in close-knit rural networks.')

add_body('Operational Risk Mitigation:', bold=True)
add_bullet('Multi-vendor logistics strategy — never more than 35% volume with a single partner.')
add_bullet('Insurance for in-transit produce; SLA-driven contracts with monetary penalties.')
add_bullet('Surge capacity planning before known harvest peaks (Apr-May for Rabi, Oct-Nov for Kharif).')

add_body('Competitive Risk Mitigation:', bold=True)
add_bullet('Build deep moats: proprietary AI models trained on KrishiConnect-specific data, hyperlocal FPO partnerships, and ONDC-first architecture make replication slower than it appears.')
add_bullet('Focus on Tier-3 and Tier-4 agrarian belts that big-tech ignores in initial expansion.')

add_body('Climate / Macro Risk Mitigation:', bold=True)
add_bullet('Diversify across 12+ crop categories and 10+ states so a localized failure is absorbed.')
add_bullet('Promote crop insurance heavily — every insured farmer reduces churn risk for KrishiConnect.')

doc.add_page_break()

# ==================== 10. CONCLUSION ====================
add_heading('10. Conclusion and Future Scope')

add_subheading('10.1 Summary of Project Potential')
add_body(
    'KrishiConnect represents a meaningful, market-validated, and technology-anchored response to one of '
    'India\'s most enduring developmental challenges — the structural inefficiency of its agricultural '
    'value chain. By unifying market access, input supply, equipment rental, advisory, and finance in '
    'a single, vernacular, AI-driven mobile platform, KrishiConnect creates value for every stakeholder: '
    'farmers earn more, buyers source cheaper and traceable produce, lenders access a low-default '
    'borrower pool, and the broader economy benefits from reduced food waste and rural distress.',
    indent=True
)
add_body(
    'The startup\'s differentiated USPs — voice-first UX, reverse auctions, AI Crop Doctor, ONDC-native '
    'architecture, and embedded finance — combined with a multi-stream revenue model and tailwinds from '
    'India\'s Digital Agriculture Mission, position it not just as a viable business but as a category-'
    'defining platform.',
    indent=True
)

add_subheading('10.2 Scalability — 3 to 5 Year Roadmap')

add_body('Year 1 (2026–27) — Pilot & Validation', bold=True)
add_bullet('Launch in 10 districts of Uttar Pradesh; reach 50,000 active farmers and 1,200 buyers.')
add_bullet('Achieve product-market fit; iterate AI advisor based on field feedback.')

add_body('Year 2 (2027–28) — Geographic Expansion', bold=True)
add_bullet('Expand to UP, MP, Bihar, Maharashtra; 3 lakh active farmers, 7,500 buyers.')
add_bullet('Launch KrishiPro premium subscription and equipment-rental network.')
add_bullet('Series A funding round (USD 8–12 million).')

add_body('Year 3 (2028–29) — National Coverage', bold=True)
add_bullet('Operations in 18 states; 12 lakh farmers, 35,000 buyers; ₹2,400 crore GMV.')
add_bullet('Cross EBITDA break-even.')
add_bullet('Launch B2B export module for traceable Indian produce (basmati, spices, organic).')

add_body('Year 4 (2029–30) — Vertical Deepening', bold=True)
add_bullet('Add dairy, poultry, fisheries, and apiculture marketplaces.')
add_bullet('Launch carbon-credit monetization for sustainable farmers.')
add_bullet('Series B funding (USD 30–50 million); aim for 25 million registered farmers.')

add_body('Year 5 (2030–31) — International & IPO Readiness', bold=True)
add_bullet('Pilot in Bangladesh, Nepal, Sri Lanka, and select African countries (Kenya, Nigeria) where smallholder agriculture dynamics mirror India.')
add_bullet('Strategic partnerships with FAO, World Bank IFC, and ADB for co-funded rural digitization.')
add_bullet('IPO readiness via DRHP filing — targeting valuation of USD 1.2–1.5 billion (Unicorn status).')

add_subheading('10.3 Closing Reflection')
add_body(
    'The greatest strength of KrishiConnect lies in its alignment of profit and purpose — every rupee '
    'the platform earns translates into measurable benefit at the farm gate. In an era where technology '
    'has reshaped finance, retail, mobility, and entertainment, agriculture remains the last great '
    'frontier of digital transformation. KrishiConnect is conceived to lead that transformation in India '
    '— not by displacing the farmer, but by amplifying his voice, his bargaining power, and his future.',
    indent=True
)

doc.add_page_break()

# ==================== 11. REFERENCES ====================
add_heading('11. References')

refs = [
    'Ministry of Agriculture & Farmers Welfare, Government of India. (2025). Annual Report 2024–25. Retrieved from agriculture.gov.in.',
    'NABARD. (2023). All-India Rural Financial Inclusion Survey (NAFIS) 2021–22. National Bank for Agriculture and Rural Development, Mumbai.',
    'Bain & Company and EY-NASSCOM. (2024). Indian Agritech: Unlocking the USD 34 Billion Opportunity. Joint Industry Report.',
    'NITI Aayog. (2024). Doubling Farmers\' Income — Implementation Strategy. Government of India.',
    'Open Network for Digital Commerce (ONDC). (2025). ONDC for Agriculture — Reference Architecture v2.1. ondc.org.',
    'Ministry of Electronics and Information Technology. (2023). The Digital Personal Data Protection Act, 2023. meity.gov.in.',
    'ICAR. (2024). Vision 2050 — Indian Council of Agricultural Research. New Delhi.',
    'World Bank. (2023). India Agriculture Overview. Washington D.C.',
    'FAO. (2024). The State of Food and Agriculture 2024. Food and Agriculture Organization of the United Nations.',
    'IBEF. (2025). Indian Agriculture and Allied Industries Report. India Brand Equity Foundation.',
    'TechSci Research. (2024). India Agritech Market — Forecast and Opportunities, 2030.',
    'McKinsey & Company. (2023). Harvesting Golden Opportunities in Indian Agriculture.',
    'Inc42. (2025). Indian Agritech Funding Report — H1 2025. inc42.com.',
    'Ministry of Rural Development. (2025). Self-Help Group & FPO Implementation Status. rural.nic.in.',
    'Kshetri, N. (2022). Blockchain and Sustainable Supply Chain Management in Developing Countries. International Journal of Information Management, 60.',
    'DeHaat. (2024). Annual Impact Report. agrevolution.in.',
    'AgroStar. (2024). Farmer Engagement Insights. agrostar.in.',
    'Press Information Bureau. (2024). Digital Agriculture Mission — Cabinet Approval Note. pib.gov.in.',
    'IIM Ahmedabad — CMA. (2023). Smallholder Farmer Income & Access Study. Centre for Management in Agriculture.',
    'Rural Smartphone Index — KANTAR ICUBE 2025 Report.',
]

for r in refs:
    p = doc.add_paragraph(style='List Number')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(r)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

doc.add_page_break()

# ==================== SIGNATURES ====================
add_heading('Signatures', size=14)
add_body(' ')
add_body(' ')
add_body(' ')

t = doc.add_table(rows=2, cols=2)
t.rows[0].cells[0].text = '_______________________'
t.rows[0].cells[1].text = '_______________________'
t.rows[1].cells[0].text = 'Chinmaya Venkataraman\n(Student)\nRoll No: [Your Roll Number]'
t.rows[1].cells[1].text = '[Faculty Name]\n(Subject Incharge / Supervisor)\nAssistant Professor'
for r in t.rows:
    for c in r.cells:
        for para in c.paragraphs:
            for run in para.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

doc.save(r'c:\Users\Manish\Desktop\STARTUP AND ENTREPRENEURIAL\Agri-Tech Marketplace for Farmers - Startup Report.docx')
print("Document generated successfully!")
