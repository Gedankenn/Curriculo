#!/usr/bin/env python3
"""Build Fabio Slika Stella's CV (EN and PT) as DOCX, then PDF via LibreOffice.

Usage: python3 src/build_cv.py   (writes the PDFs to the repo root)
Requires: python-docx, LibreOffice (soffice), Noto Serif font.
"""
import os, subprocess, sys, tempfile
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn

DARK = RGBColor(0x1F, 0x2A, 0x36)
MUTED = RGBColor(0x2C, 0x3A, 0x4A)
FONT = "Noto Serif"
SEP = "  |  "

CV = {
 "en": {
  "file": "Curriculo_Fabio_Stella_EN.pdf",
  "headline": "Senior Embedded Linux and Firmware Engineer | C/C++ | Yocto | Kernel Drivers | Bare-Metal",
  "contact": ["Pato Branco, PR, Brazil", "fabioslikastella@gmail.com", "github.com/Gedankenn", "Portuguese, English"],
  "h_summary": "Professional Summary",
  "summary": "Computer Engineer shipping firmware and Embedded Linux on production devices (electronic security, railway, automotive). Owns Yocto (BitBake), Device Tree (DTS/DTB), Linux kernel drivers/BSP and bare-metal Embedded C/C++ on ARM Cortex-M (MSP, STM32, ESP32): bootloader, OTA, cryptography, MQTT, Ethernet/TCP-IP, CAN, RS-485, I2C, UART. Works in cross-functional teams; listens closely; treats code review as engineering, not conflict.",
  "h_skills": "Technical Skills",
  "skills": [
   ("Languages: ", "C, C++, Embedded C, Python, MATLAB, Bash, Assembly"),
   ("Embedded Linux: ", "Yocto Project, BitBake, OpenEmbedded, Device Tree, DTS, DTB, Linux kernel, kernel drivers, BSP, OpenWrt, GCC"),
   ("Firmware: ", "bare-metal, bootloader, OTA, board bring-up, ARM Cortex-M, STM32, ESP32, MSP"),
   ("Connectivity / security: ", "CAN, RS-485, I2C, UART, Ethernet, TCP/IP, MQTT, 4G, cryptography, GDB, Git, CMake, MATLAB/Simulink, Typhoon HIL"),
  ],
  "h_exp": "Professional Experience",
  "exp": [
   ("Control iD  —  Sao Paulo, Brazil", "Systems Developer II  |  2026 — Present", [
     "Owned production Embedded Linux images (Yocto/BitBake, Device Tree DTS/DTB, Linux kernel drivers) for access-control devices.",
     "Developed bare-metal Embedded C/C++ on 3 ARM Cortex-M families (MSP, STM32, ESP32): bootloader, OTA and board bring-up.",
     "Integrated cryptography, Ethernet, TCP/IP, MQTT and web services; led code review and GDB debug on the R&D team."]),
   ("Progress Rail — Caterpillar", "Embedded Systems Engineer  |  2024", [
     "Shipped locomotive-controller firmware in C/C++ with CAN, RS-485 and 4G Ethernet via cellular modem (MSP).",
     "Designed data architecture, persistent storage and microcontroller memory map for field operation."]),
   ("Zeentech / Stellantis", "Software Engineer  |  2024", [
     "Built MATLAB/Simulink vehicle models and validated automotive embedded code from simulation to bench."]),
   ("InoBram", "Embedded Systems Developer  |  2022 — 2024", [
     "Wrote STM32 bare-metal firmware in C across a 2-year industrial product cycle, including bootloader, OTA and bring-up."]),
   ("IAV do Brasil", "Test Engineer  |  2022", [
     "Performed white-box HIL validation of automotive electronics with Typhoon HIL."]),
   ("Xpert Automacao", "R&D Intern  |  2021 — 2022", [
     "Implemented ESP32/STM32 bare-metal firmware; deployed OpenWrt on a commercial router adopted as the product network backbone."]),
  ],
  "h_proj": "Selected Projects",
  "proj": [
   "Built ESP32-FOC — Field-Oriented Control for BLDC motors in C. github.com/Gedankenn/ESP32-FOC",
   "Delivered Strata — local C++ LLM inference engine (OpenAI/Anthropic API). github.com/Gedankenn/Strata"],
  "h_edu": "Education",
  "edu": ("B.S. Computer Engineering", "  |  UTFPR  |  2014 — 2022  |  FOC research 2020—2022  |  Bare-Metal STM32, Udemy 2023"),
  "h_lang": "Languages",
  "lang": "Portuguese (native), English (advanced)",
 },
 "pt": {
  "file": "Curriculo_Fabio_Stella_PT.pdf",
  "headline": "Engenheiro Sênior de Linux Embarcado e Firmware | C/C++ | Yocto | Drivers de Kernel | Bare-Metal",
  "contact": ["Pato Branco, PR, Brasil", "fabioslikastella@gmail.com", "github.com/Gedankenn", "Português, Inglês"],
  "h_summary": "Resumo Profissional",
  "summary": "Engenheiro da Computação que entrega firmware e Linux embarcado em dispositivos em produção (segurança eletrônica, ferroviário, automotivo). Atua ponta a ponta: Yocto (BitBake), Device Tree (DTS/DTB), drivers de kernel Linux/BSP e C/C++ embarcado bare-metal em ARM Cortex-M (MSP, STM32, ESP32): bootloader, OTA, criptografia, MQTT, Ethernet/TCP-IP, CAN, RS-485, I2C, UART. Trabalha bem em time; escuta com precisão; trata revisão de código como engenharia, não como conflito.",
  "h_skills": "Competências Técnicas",
  "skills": [
   ("Linguagens: ", "C, C++, Embedded C, Python, MATLAB, Bash, Assembly"),
   ("Linux embarcado: ", "Yocto Project, BitBake, OpenEmbedded, Device Tree, DTS, DTB, kernel Linux, drivers de kernel, BSP, OpenWrt, GCC"),
   ("Firmware: ", "bare-metal, bootloader, OTA, bring-up de placas, ARM Cortex-M, STM32, ESP32, MSP"),
   ("Conectividade / segurança: ", "CAN, RS-485, I2C, UART, Ethernet, TCP/IP, MQTT, 4G, criptografia, GDB, Git, CMake, MATLAB/Simulink, Typhoon HIL"),
  ],
  "h_exp": "Experiência Profissional",
  "exp": [
   ("Control iD  —  São Paulo, Brasil", "Desenvolvedor de Sistemas II  |  2026 — Atual", [
     "Responsável pelas imagens de Linux embarcado de produção (Yocto/BitBake, Device Tree DTS/DTB, drivers de kernel) para controle de acesso.",
     "Desenvolveu firmware bare-metal em C/C++ embarcado para 3 famílias ARM Cortex-M (MSP, STM32, ESP32): bootloader, OTA e bring-up.",
     "Integrou criptografia, Ethernet, TCP/IP, MQTT e serviços web; liderou revisões de código e depuração com GDB no time de P&D."]),
   ("Progress Rail — Caterpillar", "Engenheiro de Sistemas Embarcados  |  2024", [
     "Entregou firmware de controladoras de locomotivas em C/C++ com CAN, RS-485 e 4G Ethernet via modem celular (MSP).",
     "Projetou a arquitetura de dados, o armazenamento persistente e o mapa de memória do microcontrolador."]),
   ("Zeentech / Stellantis", "Engenheiro de Software  |  2024", [
     "Desenvolveu modelos em MATLAB/Simulink e validou código embarcado automotivo da simulação à bancada."]),
   ("InoBram", "Desenvolvedor de Sistemas Embarcados  |  2022 — 2024", [
     "Programou firmware bare-metal para STM32 em C ao longo de um ciclo de produto de 2 anos, com bootloader, OTA e bring-up."]),
   ("IAV do Brasil", "Analista de Testes  |  2022", [
     "Realizou validação caixa-branca em HIL de eletrônica veicular com Typhoon HIL."]),
   ("Xpert Automacao", "Estágio em P&D  |  2021 — 2022", [
     "Implementou firmware bare-metal para ESP32/STM32; implantou OpenWrt em roteador comercial, adotado como backbone de rede do produto."]),
  ],
  "h_proj": "Projetos em Destaque",
  "proj": [
   "Desenvolveu o ESP32-FOC — controle por orientação de campo (FOC) para motores BLDC em C. github.com/Gedankenn/ESP32-FOC",
   "Criou o Strata — motor de inferência LLM em C++ (API OpenAI/Anthropic). github.com/Gedankenn/Strata"],
  "h_edu": "Formação",
  "edu": ("Engenharia da Computação", "  |  UTFPR  |  2014 — 2022  |  pesquisa FOC 2020—2022  |  Bare-Metal STM32, Udemy 2023"),
  "h_lang": "Idiomas",
  "lang": "Português (nativo), Inglês (avançado)",
 },
}

def para(doc, before=0, after=0, indent=None, hanging=None, line=12.7):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line)
    if indent is not None: pf.left_indent = Pt(indent)
    if hanging is not None: pf.first_line_indent = Pt(-hanging)
    return p

def run(p, text, size=10, bold=False, italic=False, color=DARK):
    r = p.add_run(text)
    r.font.name = FONT; r.font.size = Pt(size); r.bold = bold; r.italic = italic
    r.font.color.rgb = color
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return r

def heading(doc, text):
    run(para(doc, before=3, after=1, line=14), text, 11, bold=True)

def bullet(doc, text):
    run(para(doc, indent=11), "•\u00a0" + text)

def build(d, out_docx):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.3)
    s.top_margin = Cm(0.9); s.bottom_margin = Cm(0.9)
    st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10)
    run(para(doc, after=1, line=20), "FABIO SLIKA STELLA", 16, bold=True)
    run(para(doc, line=14), d["headline"].replace(" | ", SEP), 11, bold=True, color=MUTED)
    run(para(doc, after=1), SEP.join(d["contact"]), 10, color=MUTED)
    heading(doc, d["h_summary"]); run(para(doc), d["summary"])
    heading(doc, d["h_skills"])
    for k, v in d["skills"]:
        p = para(doc); run(p, k, bold=True); run(p, v)
    heading(doc, d["h_exp"])
    for org, role, bullets in d["exp"]:
        run(para(doc, before=1, line=13.5), org, 10.5, bold=True)
        run(para(doc), role, italic=True, color=MUTED)
        for b in bullets: bullet(doc, b)
    heading(doc, d["h_proj"])
    for b in d["proj"]: bullet(doc, b)
    heading(doc, d["h_edu"])
    p = para(doc); run(p, d["edu"][0], bold=True); run(p, d["edu"][1])
    heading(doc, d["h_lang"]); run(para(doc), d["lang"])
    doc.core_properties.author = "Fabio Slika Stella"
    doc.core_properties.title = "Fabio Slika Stella — CV"
    doc.save(out_docx)

if __name__ == "__main__":
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    tmp = tempfile.mkdtemp()
    for lang, d in CV.items():
        stem = os.path.splitext(d["file"])[0]
        docx = os.path.join(tmp, stem + ".docx")
        build(d, docx)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", root, docx], check=True)
        print("built", os.path.join(root, d["file"]))
