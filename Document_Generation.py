# Import necessary classes from the python-docx library
from docx import Document
from docx.shared import Inches, Pt # For setting margins, font sizes, etc. (optional)
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT # For text alignment (optional)

# --- Content from the provided resume ---
# (Extracted manually for clarity in this example)

name = "Arthur Santos Marinho de Souza"
contact_info = "arthursmds@gmail.com | (81) 98605-6957 | https://www.linkedin.com/in/asmdes/"

awards_certs_title = "Prêmios e Certificações"
awards_certs_details = [
    "Medalhas: Ouro - Olimpíada Brasileira de Astronomia (2019/2022), Prata - Olimpíada Nacional de Ciências (2022), Prata - Olimpíada Pernambucana de Astronomia (2022), Menção Honrosa - Olimpíada Pernambucana de Física (2022)",
    "Inglês: ECPE (Pass), Duolingo English Test (140/160), SAT (1470/1600)"
]

experience_title = "Experiência"
experience_details = [
    {
        "title": "Planejamento de Supply Chain, Praso – Recife, Pernambuco",
        "duration": "Agosto 2024 – Presente",
        "points": [
            "Reconstruí as ferramentas de previsão de demanda, melhorando a acurácia do nosso planejamento de compras e o risco de estocagem em excesso",
            "Reduzi as perdas por shelf (validade) com um sistema de prevenção de perdas que utiliza prazos de validade e estocagem em excesso, combinados com o sistema de previsão de demanda",
            "Construi uma ferramenta de precificação de produtos, garantindo uma rentabilidade adequada e alinhada com os objetivos da empresa"
        ]
    }
]

organization_title = "Organização"
organization_details = [
     {
        "title": "Membro, UFPE Finance (Liga Acadêmica) – Recife, Pernambuco",
        "duration": "Abril 2024 – Presente",
        "points": [
            "Trainee: Desenvolvi scripts python para realização do valuation intrínseco da Arezzo, analisando fluxo de caixa e balanço patrimonial de forma automatizada, tendo como base teórica os princípios de Aswath Damodaran, além de avaliar o modelo de negócios da empresa.",
            "Itaú Quant: Desenvolvi e validei uma estratégia sistemática de investimentos baseada no indicador KST (Known-Sure Thing), focada na ação TSMC, utilizando dados históricos via API para capturar movimentos de curto prazo e superar o benchmark S&P 500",
            "Constellation Challenge: construí uma análise dos setores de E-Commerce e Fintech na América Latina, para realização de um benchmark para o Mercado Livre"
        ]
    }
]

volunteer_title = "Trabalho Voluntário"
volunteer_details = [
    {
        "title": "Professor de Matemática, Voluntariado Estudantil Marista – Recife, Pernambuco",
        "duration": "Março 2022 – Presente",
        "points": [
            "Ensinei Matemática básica, como as quatro operações, para crianças carentes em idade escolar",
            "Desenvolvi de atividades lúdicas para aprendizado de habilidades socioemocionais",
            "Interagi com as crianças durante os intervalos"
        ]
    }
]

projects_title = "Projetos"
projects_details = [
    {
        "title": "Stock selection performance - Dashboard",
        "link": "https://www.upwork.com/freelancers/~01dd196048866f8814",
        "points": [
            "Desenvolvi um dashboard dinâmico integrado a planilhas online e sites financeiros, permitindo monitoramento em tempo real de KPIs de desempenho de ações, com feedback máximo do cliente.",
            "Implementei um módulo de backtesting utilizando dados de 2023, identificando padrões de desempenho e otimizando estratégias de seleção de ações.",
            "Criei uma interface intuitiva com filtros personalizados e documentei todo o sistema, garantindo usabilidade e suporte pós-implementação eficiente."
        ]
    }
]

education_title = "Formação Acadêmica"
education_details = [
    "Universidade Federal de Pernambuco (UFPE) – Bacharelado em Sistemas de Informação\t  2023.1 - 2026.2"
]

# --- Create the DOCX document ---

# Create an instance of a Document
doc = Document()

# --- Add Name and Contact Info ---
# Add the name as a main heading (or a styled paragraph)
name_paragraph = doc.add_paragraph()
name_run = name_paragraph.add_run(name)
name_run.bold = True
name_run.font.size = Pt(16) # Optional: Set font size
name_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # Optional: Center align

# Add contact info
contact_paragraph = doc.add_paragraph(contact_info)
contact_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # Optional: Center align

# --- Add Awards and Certifications ---
doc.add_heading(awards_certs_title, level=1) # level=1 for a main section heading
for item in awards_certs_details:
    doc.add_paragraph(item, style='List Bullet')

# --- Add Experience ---
doc.add_heading(experience_title, level=1)
for job in experience_details:
    # Add job title and duration
    p_title = doc.add_paragraph()
    p_title.add_run(job['title']).bold = True
    p_title.add_run(f"\t {job['duration']}") # Add duration, maybe adjust spacing/tab

    # Add bullet points for responsibilities/achievements
    for point in job['points']:
        doc.add_paragraph(point, style='List Bullet')

# --- Add Organization ---
doc.add_heading(organization_title, level=1)
for org in organization_details:
    # Add organization title and duration
    p_title = doc.add_paragraph()
    p_title.add_run(org['title']).bold = True
    p_title.add_run(f"\t {org['duration']}")

    # Add bullet points
    for point in org['points']:
        # Check if the point itself contains a sub-heading like "Trainee:"
        if ":" in point:
             parts = point.split(":", 1)
             p_bullet = doc.add_paragraph(style='List Bullet')
             p_bullet.add_run(parts[0] + ":").bold = True
             p_bullet.add_run(parts[1])
        else:
            doc.add_paragraph(point, style='List Bullet')


# --- Add Volunteer Work ---
doc.add_heading(volunteer_title, level=1)
for volunteer in volunteer_details:
    # Add title and duration
    p_title = doc.add_paragraph()
    p_title.add_run(volunteer['title']).bold = True
    p_title.add_run(f"\t {volunteer['duration']}")

    # Add bullet points
    for point in volunteer['points']:
        doc.add_paragraph(point, style='List Bullet')

# --- Add Projects ---
doc.add_heading(projects_title, level=1)
for project in projects_details:
    # Add project title and link
    p_title = doc.add_paragraph()
    p_title.add_run(project['title']).bold = True
    if 'link' in project: # Add link if present
         p_title.add_run(f"\t{project['link']}") # Consider making this a hyperlink if needed

    # Add bullet points
    for point in project['points']:
        doc.add_paragraph(point, style='List Bullet')

# --- Add Education ---
doc.add_heading(education_title, level=1)
for edu in education_details:
    doc.add_paragraph(edu) # Add education details as a simple paragraph

# --- Save the document ---
file_name = 'Arthur_Souza_Resume.docx'
try:
    doc.save(file_name)
    print(f"Document '{file_name}' created successfully.")
except Exception as e:
    print(f"Error saving document: {e}")

