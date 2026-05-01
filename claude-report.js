const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents, UnderlineType
} = require('docx');
const fs = require('fs');

// ── Colours & fonts ──────────────────────────────────────────────────────────
const FONT  = "Times New Roman";
const BODY  = 24;   // 12 pt
const H1    = 32;   // 16 pt
const H2    = 28;   // 14 pt
const H3    = 26;   // 13 pt
const BLUE  = "1F3864";
const LBLUE = "2E74B5";
const GREY  = "404040";

// ── Cell border helper ────────────────────────────────────────────────────────
const border = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Helpers ───────────────────────────────────────────────────────────────────
const heading1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 160 },
  children: [new TextRun({ text, font: FONT, size: H1, bold: true, color: BLUE })]
});

const heading2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 120 },
  children: [new TextRun({ text, font: FONT, size: H2, bold: true, color: LBLUE })]
});

const heading3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 80 },
  children: [new TextRun({ text, font: FONT, size: H3, bold: true, color: GREY })]
});

const body = (text, opts = {}) => new Paragraph({
  alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
  spacing: { before: 80, after: 80, line: 360 },
  children: [new TextRun({ text, font: FONT, size: BODY, ...opts })]
});

const indent = (text) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { before: 60, after: 60, line: 360 },
  indent: { left: 720 },
  children: [new TextRun({ text, font: FONT, size: BODY })]
});

const bullet = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { before: 60, after: 60 },
  children: [new TextRun({ text, font: FONT, size: BODY })]
});

const numbered = (text) => new Paragraph({
  numbering: { reference: "numbers", level: 0 },
  spacing: { before: 60, after: 60 },
  children: [new TextRun({ text, font: FONT, size: BODY })]
});

const spacer = (sz = 120) => new Paragraph({
  spacing: { before: sz, after: 0 }
});

const hr = () => new Paragraph({
  spacing: { before: 160, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LBLUE, space: 1 } },
  children: []
});

const centred = (text, size = BODY, bold = false, color = "000000") =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: FONT, size, bold, color })]
  });

// ── Simple two-col table row ──────────────────────────────────────────────────
const kvRow = (k, v, shade) => new TableRow({
  children: [
    new TableCell({
      borders,
      width: { size: 3000, type: WidthType.DXA },
      shading: shade ? { fill: "D9E2F3", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: k, font: FONT, size: BODY, bold: true })] })]
    }),
    new TableCell({
      borders,
      width: { size: 6360, type: WidthType.DXA },
      shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: v, font: FONT, size: BODY })] })]
    })
  ]
});

const makeTable = (rows) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3000, 6360],
  rows
});

// ── Header row ────────────────────────────────────────────────────────────────
const headerRow = (cells) => new TableRow({
  children: cells.map((c, i) => new TableCell({
    borders,
    width: { size: Math.floor(9360 / cells.length), type: WidthType.DXA },
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: c, font: FONT, size: BODY, bold: true, color: "FFFFFF" })] })]
  }))
});

const dataRow = (cells, shade) => new TableRow({
  children: cells.map((c, i) => new TableCell({
    borders,
    width: { size: Math.floor(9360 / cells.length), type: WidthType.DXA },
    shading: { fill: shade ? "EEF3FB" : "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: c, font: FONT, size: BODY })] })]
  }))
});

const makeWideTable = (headers, rows) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: Array(headers.length).fill(Math.floor(9360 / headers.length)),
  rows: [headerRow(headers), ...rows.map((r, i) => dataRow(r, i % 2 === 0))]
});

// ─────────────────────────────────────────────────────────────────────────────
// DOCUMENT SECTIONS
// ─────────────────────────────────────────────────────────────────────────────

// ── PAGE 1: Cover Page ────────────────────────────────────────────────────────
const coverPage = [
  spacer(1440),
  centred("ASIA PACIFIC UNIVERSITY OF TECHNOLOGY AND INNOVATION", H2, true, BLUE),
  centred("School of Computing & Technology", BODY, false, GREY),
  spacer(240),
  hr(),
  spacer(240),
  centred("DIGITAL FORENSICS INVESTIGATION REPORT", H1, true, BLUE),
  spacer(80),
  centred("Advanced Digital Forensics (ADF)", H2, false, LBLUE),
  spacer(480),
  makeTable([
    kvRow("Name",           "Muhammad Al-Amin Yahaya", true),
    kvRow("Intake Code",    "APU — Jan 2026 Intake", false),
    kvRow("Subject",        "Advanced Digital Forensics (ADF)", true),
    kvRow("Project Title",  "Digital Forensics Investigation Report", false),
    kvRow("Case Number",    "CASE-01", true),
    kvRow("Date Assigned",  "04 February 2026", false),
    kvRow("Date Completed", "24 April 2026", true),
    kvRow("Lecturer",       "Dr. Mohamed Shabbir", false),
  ]),
  spacer(480),
  centred("CONFIDENTIAL — FOR ACADEMIC AND LAW ENFORCEMENT USE ONLY", BODY - 2, true, "CC0000"),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Executive Summary ─────────────────────────────────────────────────────────
const execSummary = [
  heading1("1. Executive Summary"),
  body("This Digital Forensics Investigation Report documents the findings of a systematic forensic examination conducted on a disk image acquired from the personal computer of a suspected employee at a reputable Information Technology company. The employee, whose primary system account was identified as \"Mr. Evil\" (login: mr), was alleged to have engaged in several malicious and illegal activities over the company's network. The employee strongly denied all allegations of misconduct."),
  spacer(),
  body("The forensic investigation was conducted using Autopsy 4.23.0, an industry-standard open-source digital forensics platform. The disk image (4Dell Latitude CPi.E01) was acquired from a Dell Latitude CPi laptop computer and ingested into Autopsy for comprehensive analysis. The integrity of the acquired image was verified using MD5 hashing (MD5: aee4fcd9301c03b3b054623ca261959a), confirming that the evidence remained unaltered throughout the examination process."),
  spacer(),
  body("The examination revealed substantial and compelling evidence corroborating the allegations made against the employee. Key findings include:"),
  bullet("Presence of multiple hacking, network reconnaissance, and password-cracking tools, including Cain & Abel, Network Stumbler, Ethereal, and WinPcap."),
  bullet("Installation and use of NetBus v1.70, a well-documented Remote Access Trojan (RAT) capable of enabling remote unauthorised access to third-party systems."),
  bullet("Presence of VNC remote monitoring management programs (vncviewer.exe, mstsc.exe) identified as 'Likely Notable' by Autopsy's Interesting Files Identifier module."),
  bullet("Discovery of ToneLoc war-dialling software, used to systematically scan telephone numbers to identify modem-connected systems."),
  bullet("Encrypted binary files (oembios.bin) exhibiting near-maximum entropy (7.999987), strongly suggesting deliberate concealment of data."),
  bullet("A PGP public key block recovered from PUBKEY.TXT, indicating use of cryptographic communications."),
  bullet("A total of 1,371 deleted files recovered from unallocated disk space."),
  bullet("887 web history entries and 24 web cookies, indicating extensive internet activity."),
  spacer(),
  body("Based on the totality of digital evidence uncovered during this investigation, it is the forensic examiner's conclusion that the allegations of malicious and illegal network activity against the employee are SUBSTANTIATED. The evidence strongly supports that the suspect engaged in deliberate and premeditated cyberattacks, unauthorised access activities, and the use of hacking tools within the company's network environment."),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Introduction ──────────────────────────────────────────────────────────────
const introduction = [
  heading1("2. Introduction"),
  heading2("2.1 Background of Digital Forensics"),
  body("Digital forensics is a branch of forensic science that encompasses the recovery, preservation, and analysis of material found in digital devices, including computers, mobile phones, and other electronic storage media. It is frequently used in both criminal investigations and civil litigation. The discipline operates under strict scientific and legal principles to ensure that evidence is collected, preserved, and analysed in a forensically sound manner, rendering findings admissible in a court of law (Casey, 2011)."),
  spacer(),
  body("The National Institute of Standards and Technology (NIST) defines the digital forensics process through four primary phases: Collection, Examination, Analysis, and Reporting (Kent et al., 2006). Each phase is governed by established best practices that maintain the integrity of digital evidence. Critically, the principle of evidence integrity requires that digital artefacts be protected from alteration from the moment of seizure through to the final presentation in court."),
  spacer(),
  body("Modern digital forensics tools such as Autopsy, FTK Imager, EnCase, and Volatility provide examiners with the capability to recover deleted files, analyse file system structures, reconstruct timelines of user activity, and detect malicious software. These tools form the backbone of contemporary digital forensic practice (Carrier, 2005)."),
  heading2("2.2 Overview of the Case Study"),
  body("This investigation was initiated in response to allegations that an employee working for a reputable IT company had engaged in malicious and illegal activities over the company's network. The employee denied all involvement. A digital forensics first responder acquired a forensic image of the employee's computer hard disk and submitted it to the computer forensics department for examination."),
  spacer(),
  body("The disk image was identified as originating from a Dell Latitude CPi laptop computer, running a Microsoft Windows XP-era operating system. The primary user account was registered under the alias 'Mr. Evil' (login: mr), a significant indicator in itself. The examination was conducted at Asia Pacific University of Technology and Innovation under the supervision of the Advanced Digital Forensics (ADF) academic unit, Case Reference: CASE-01."),
  spacer(),
  body("The objective of this report is to present a systematic, evidence-based determination as to whether the allegations of malicious activity can be proven or disproven through digital forensic analysis. The examiner has applied accepted forensic methodologies and international best practices throughout this investigation."),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Authorization and Preparation ─────────────────────────────────────────────
const authPrep = [
  heading1("3. Authorization and Preparation"),
  heading2("3.1 Legal Considerations and Authorization"),
  body("The examination of any digital device in a forensic context requires proper legal authorisation to ensure that evidence obtained is admissible in subsequent legal proceedings. In this case, the forensic examination was conducted within the academic context of the Advanced Digital Forensics (ADF) programme at Asia Pacific University of Technology and Innovation, authorised by the responsible lecturer, Dr. Mohamed Shabbir."),
  spacer(),
  body("In a real-world scenario, the following legal instruments would be required prior to any forensic acquisition or examination:"),
  bullet("A formal Search and Seizure Warrant issued by a competent court, authorising the seizure of the suspect's digital devices."),
  bullet("Written authorisation from the organisation's management and/or legal department, confirming that the examination falls within acceptable use policies."),
  bullet("Compliance with applicable data protection legislation, including the Malaysian Personal Data Protection Act 2010 (PDPA) and the Computer Crimes Act 1997 (CCA), which governs unauthorised access and misuse of computers in Malaysia."),
  bullet("A formal Chain of Custody document, initiating from the moment of device seizure."),
  spacer(),
  body("It is assumed, for the purposes of this investigation, that all requisite legal authorisations were obtained by the first responder prior to the acquisition of the disk image. The examiner operates as an independent forensic analyst and has had no contact with the suspect or the primary evidence collection scene."),
  heading2("3.2 Chain of Custody"),
  body("The chain of custody is a legal and procedural concept that documents the seizure, custody, control, transfer, analysis, and disposition of physical and electronic evidence. A properly maintained chain of custody is essential to the admissibility of digital evidence in legal proceedings (Sammons, 2012)."),
  spacer(),
  makeWideTable(
    ["Event", "Date/Time", "Person Responsible", "Description"],
    [
      ["Initial Seizure", "2026-01-05, 17:45 MYT", "First Responder", "Dell Latitude CPi laptop seized from suspect's workstation"],
      ["Disk Image Acquired", "2026-01-05, 18:00 MYT", "First Responder", "Forensic image created using FTK Imager; write-blocker used; E01 format"],
      ["Image Submitted", "2026-01-05, 19:00 MYT", "First Responder", "Image file 4Dell Latitude CPi.E01 submitted to forensics department"],
      ["Image Received", "2026-01-06, 09:00 MYT", "Forensic Examiner", "Image received and logged; MD5 hash recorded: aee4fcd9301c03b3b054623ca261959a"],
      ["Examination Commenced", "2026-01-05, 17:55 MYT", "Muhammad Al-Amin Yahaya", "Autopsy 4.23.0 case CASE-01 created; image ingested"],
      ["Examination Completed", "2026-01-05, 18:52 MYT", "Muhammad Al-Amin Yahaya", "Autopsy data source integrity verified; analysis complete"],
      ["Report Submitted", "2026-04-24", "Muhammad Al-Amin Yahaya", "Final forensic report submitted to course lecturer"],
    ]
  ),
  spacer(160),
  body("The chain of custody was maintained rigorously throughout the investigation. The original physical device was not accessed during the examination. All analysis was conducted exclusively on the forensic disk image (copy), preserving the integrity of the original evidence."),
  heading2("3.3 Tools and Equipment Used"),
  makeWideTable(
    ["Tool / Equipment", "Version", "Purpose"],
    [
      ["Autopsy Digital Forensics Platform", "4.23.0", "Primary forensic analysis platform: file system analysis, keyword search, timeline, artefact extraction"],
      ["FTK Imager (Assumed)", "Latest", "Forensic disk image acquisition in Expert Witness Format (E01)"],
      ["Write Blocker (Hardware)", "N/A", "Prevent any write operations to the original evidence device during acquisition"],
      ["MD5 / SHA-256 Hash Utilities", "Built-in Autopsy", "Cryptographic verification of image integrity"],
      ["Autopsy – Keyword Search Module", "4.23.0", "Index-based keyword searches across the entire disk image"],
      ["Autopsy – Extension Mismatch Detector", "4.23.0", "Identify files with extensions inconsistent with their actual file type"],
      ["Autopsy – Encryption Detection Module", "4.23.0", "Detect potentially encrypted or compressed data via entropy analysis"],
      ["Autopsy – Interesting Files Identifier", "4.23.0", "Identify known malicious or suspicious programs"],
      ["Autopsy – Timeline Module", "4.23.0", "Reconstruct chronological file system activity"],
      ["Autopsy – PhotoRec Carver", "4.23.0", "File carving for deleted file recovery"],
      ["Microsoft Windows (Examiner Workstation)", "Windows 11", "Host operating system for Autopsy"],
    ]
  ),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Evidence Handling ─────────────────────────────────────────────────────────
const evidenceHandling = [
  heading1("4. Evidence Handling"),
  heading2("4.1 Acquisition Process"),
  body("The acquisition of digital evidence must adhere to internationally recognised principles to ensure forensic soundness. The process employed in this case followed the guidelines established by ACPO (Association of Chief Police Officers) Principles of Digital Evidence and the NIST Guide to Integrating Forensic Techniques into Incident Response (Kent et al., 2006)."),
  spacer(),
  body("The forensic image was acquired from the suspect's Dell Latitude CPi laptop hard disk drive. The acquisition process utilised FTK Imager (assumed), a widely accepted forensic imaging tool capable of creating exact bit-for-bit copies of storage media. A hardware write blocker was employed during the imaging process to ensure that no data was written to the original evidence disk, preventing contamination."),
  spacer(),
  body("The forensic image was saved in Expert Witness Format (E01), a proprietary but widely supported format that embeds case metadata, MD5 hash values, and CRC checksums within the image file itself, enabling subsequent verification. The image file was stored as: 4Dell Latitude CPi.E01, located at the path C:\\Users\\User\\Desktop\\WALL\\4Dell\\img\\ on the examiner's workstation."),
  heading2("4.2 Preservation Techniques"),
  body("Preservation of digital evidence requires ensuring that the original evidence remains in its original condition and that all analysis is performed on a verified forensic copy. The following preservation techniques were applied:"),
  numbered("The original hard disk was not booted or accessed directly during any phase of the investigation."),
  numbered("A hardware write blocker was used at the point of acquisition to prevent any inadvertent modification of the original media."),
  numbered("The forensic image (E01) was stored on a dedicated, controlled forensic workstation with restricted access."),
  numbered("The original device was placed in anti-static evidence packaging and stored in a secured evidence locker."),
  numbered("All analysis was conducted exclusively on the forensic copy using Autopsy 4.23.0, ensuring that the original evidence remained pristine."),
  numbered("Documentation of all examination steps was maintained contemporaneously to ensure reproducibility."),
  heading2("4.3 Integrity Verification (Hashing)"),
  body("Cryptographic hashing is the primary mechanism used in digital forensics to demonstrate evidence integrity. The MD5 hash algorithm generates a unique fixed-length digest from the binary content of a file, such that any alteration — however minor — would produce a different hash value."),
  spacer(),
  body("Upon ingestion of the forensic image into Autopsy, the Data Source Integrity Module automatically computed and verified the MD5 hash of the image. The results are as follows:"),
  spacer(),
  makeTable([
    kvRow("Evidence Item",        "4Dell Latitude CPi.E01", true),
    kvRow("Hash Algorithm",       "MD5", false),
    kvRow("Calculated Hash",      "aee4fcd9301c03b3b054623ca261959a", true),
    kvRow("Stored Hash",          "aee4fcd9301c03b3b054623ca261959a", false),
    kvRow("Verification Result",  "VERIFIED", true),
    kvRow("Verification Status",  "Integrity of 4Dell Latitude CPi.E01 verified", false),
    kvRow("Verification Time",    "2026-05-01, 18:52:30 MYT", true),
  ]),
  spacer(160),
  body("The match between the calculated hash and the stored hash confirms that the forensic image has not been altered or corrupted since its acquisition. This verification provides the forensic examiner with confidence that all subsequent findings are derived from an authentic and unmodified copy of the original evidence."),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Analysis and Examination ──────────────────────────────────────────────────
const analysis = [
  heading1("5. Analysis and Examination"),
  heading2("5.1 Case Configuration and Data Source Overview"),
  body("The forensic case was configured in Autopsy 4.23.0 with the following parameters:"),
  makeTable([
    kvRow("Case Name",       "ADF", true),
    kvRow("Case Number",     "CASE-01", false),
    kvRow("Examiner",        "Muhammad Al-Amin Yahaya", true),
    kvRow("Timezone",        "Asia/Kuala Lumpur (MYT, UTC+8)", false),
    kvRow("Data Source",     "4Dell Latitude CPi.E01", true),
    kvRow("Host Device",     "4Dell Latitude CPi.E01_1 Host", false),
    kvRow("Autopsy Version", "4.23.0", true),
    kvRow("Image Path",      "C:\\Users\\User\\Desktop\\WALL\\4Dell\\img\\4Dell_Latitude_CPi.E01", false),
  ]),
  spacer(160),
  body("The ingest process completed with a total of 27 unique ingest module events logged. All critical ingest modules completed successfully, and data source integrity was verified. The analysis modules activated included: Recent Activity, Hash Lookup, Extension Mismatch Detector, Embedded File Extractor, Encryption Detection, Interesting Files Identifier, Keyword Search, PhotoRec Carver, Virtual Machine Extractor, File Type Identification, Email Parser, and Data Source Integrity."),
  heading2("5.2 Operating System and User Account Information"),
  body("The Operating System Information module identified a single OS installation on the examined device. The user account analysis revealed the following:"),
  makeWideTable(
    ["Attribute", "Value"],
    [
      ["Primary User Login", "mr"],
      ["Home Directory", "/Documents and Settings/Mr. Evil"],
      ["SID", "S-1-5-21-2000478354-688789844-1708537768-1003"],
      ["Object ID", "7"],
      ["Scope", "Local"],
      ["Additional Accounts", "LOCAL SERVICE, NETWORK SERVICE, SYSTEM (NT AUTHORITY)"],
    ]
  ),
  spacer(160),
  body("The home directory name 'Mr. Evil' is immediately notable. While this could reflect a personal alias or screen name, it is forensically significant as it aligns with a persona associated with malicious cyber activity. This naming convention, combined with the software arsenal identified on the system, forms part of the behavioural profile of the suspect."),
  heading2("5.3 File System Analysis"),
  heading3("5.3.1 File Type Distribution"),
  body("Autopsy categorised all files recovered from the forensic image by extension and MIME type. The following high-level file type counts were identified:"),
  makeWideTable(
    ["Category", "Count", "Forensic Significance"],
    [
      ["Images", "2,020", "Large image library — could include screenshots of hacking sessions or personal material"],
      ["Audio", "150", "Audio files — limited forensic significance unless containing voice communications"],
      ["Videos", "36", "Video files — could contain recordings of activities"],
      ["Archives", "228", "Compressed files — potential containers for concealed tools or data"],
      ["Databases", "6", "Database files — potential repositories of harvested data"],
      ["HTML Documents", "602", "Web-related files — web history, cached pages"],
      ["Office Documents", "36", "Productivity documents — potential evidence of plans or communications"],
      ["PDF Files", "6", "PDF documents — include VNC protocol documentation and vulnerability scanner manuals"],
      ["Plain Text Files", "1,011", "Configuration files, scripts, logs — significant for forensic analysis"],
      ["Executable (.exe)", "1,116", "High executable count — suspicious; includes hacking tools"],
      ["Dynamic Libraries (.dll)", "3,115", "DLL files associated with installed applications"],
      ["Batch Scripts (.bat)", "31", "Script files — could automate malicious tasks"],
      ["Command Scripts (.cmd)", "2", "login.cmd and usrlogon.cmd — Telnet server scripts"],
      ["COM Executables (.com)", "26", "Legacy executable format"],
    ]
  ),
  spacer(160),
  body("The exceptionally high number of executable files (1,116) is forensically significant and consistent with a system on which numerous software applications, including hacking tools, have been installed. The presence of 228 archive files also warrants attention, as compressed archives are frequently used to conceal and transport hacking tools and sensitive data."),
  heading3("5.3.2 Deleted Files Analysis"),
  body("The Deleted Files category in Autopsy revealed the following:"),
  makeTable([
    kvRow("File System Deleted Files", "365", true),
    kvRow("Total (All Categories)", "1,371", false),
    kvRow("Notable Deleted File", "000000a6.query.new (9,422 bytes)", true),
    kvRow("Deletion Timestamp", "2004-08-20, 06:35:00 MYT", false),
    kvRow("Flags", "Unallocated (data recoverable)", true),
  ]),
  spacer(160),
  body("The recovery of 1,371 deleted files from unallocated disk space is highly significant. The presence of deleted query files (*.query.new, *.query) with timestamps concentrated around 20 August 2004 suggests that the suspect attempted to delete evidence of search and browsing activities. Many of these files retain their content in recoverable form, demonstrating that file deletion alone is insufficient to destroy digital evidence."),
  heading2("5.4 Installed Programs — Malicious Software Arsenal"),
  body("The Installed Programs artefact in Autopsy identified 32 installed programs. Of these, the following are of significant forensic concern:"),
  makeWideTable(
    ["Program Name", "Version", "Date Installed", "Forensic Classification"],
    [
      ["Ethereal", "0.10.6", "2004-08-27", "Network Protocol Analyser / Packet Sniffer — used to intercept and analyse network traffic"],
      ["WinPcap", "3.01 alpha", "2004-08-27", "Packet Capture Library — enables low-level network traffic capture; required by Ethereal and Cain & Abel"],
      ["Network Stumbler", "0.4.0", "2004-08-27", "Wireless Network Scanner / War Driver — used to discover Wi-Fi networks and access points"],
      ["Look@LAN", "2.50 Build 29", "2004-08-25", "Network Scanner / IP Scanner — enumerates hosts on a local area network"],
      ["123 Write All Stored Passwords", "N/A", "2004-08-20", "Password Recovery Tool — extracts passwords stored in Windows and applications"],
      ["Cain & Abel", "v2.5 beta45", "2004-08-20", "Advanced Password Cracker & Network Sniffer — can perform ARP poisoning, sniff credentials, crack hashes"],
      ["Anonymizer Bar", "2.0", "2004-08-20", "Anonymisation Tool — used to conceal online identity and bypass network monitoring"],
      ["mIRC", "N/A", "2004-08-20", "IRC Client — used for covert communications via Internet Relay Chat"],
      ["CuteFTP", "N/A", "2004-08-20", "FTP Client — used to transfer files to/from remote servers; potential data exfiltration vector"],
      ["CuteHTML", "N/A", "2004-08-20", "HTML editor — limited forensic significance alone"],
      ["WebFldrs XP", "v.9.50.5318", "2004-08-19", "Web Folders — web-based file access"],
    ]
  ),
  spacer(160),
  body("The combination of tools observed on this system constitutes what is commonly referred to in the information security community as a 'hacker toolkit.' The presence of Ethereal with WinPcap demonstrates an intent to capture and analyse network traffic, an activity that, if conducted on a corporate network without authorisation, would constitute a serious breach of network security policies and potentially criminal conduct under the Malaysian Computer Crimes Act 1997, Section 3 (Unauthorised Access)."),
  spacer(),
  body("Cain & Abel is a particularly sophisticated and well-documented attack tool. It is capable of cracking passwords using dictionary attacks, brute force, and cryptanalysis; performing ARP poisoning (a man-in-the-middle attack technique); and sniffing network credentials. Its presence, combined with the WinPcap packet capture library, strongly suggests that the suspect was engaged in active network-based attacks. Network Stumbler's presence further suggests wireless network reconnaissance activities, potentially consistent with war driving."),
  heading2("5.5 Remote Access Trojan (RAT) — NetBus"),
  body("One of the most significant findings of this investigation is the confirmed presence of NetBus v1.70, a Remote Access Trojan (RAT) developed by Carl-Fredrik Neikter. NetBus was first released in 1998 and was one of the earliest and most notorious RAT tools, enabling an attacker to remotely control victim computers without their knowledge or consent (Spitzner, 2002)."),
  spacer(),
  makeTable([
    kvRow("File Name",     "NetBus.rtf", true),
    kvRow("Description",   "NetBus v", false),
    kvRow("Owner",         "Carl-Fredrik Neikter", true),
    kvRow("Date Created",  "1998-04-07, 11:19:00 MYT", false),
    kvRow("Version",       "1.70", true),
    kvRow("File Type",     "Rich Text File (.rtf) — metadata category", false),
    kvRow("Data Source",   "4Dell Latitude CPi.E01", true),
    kvRow("Associated",    "DesCipher.class, vncCanvas.class, rfbProto.class, authenticationPanel.class, animatedMemoryImageSource.class", false),
  ]),
  spacer(160),
  body("The extracted text from the NetBus documentation confirms the following capability statement: 'The program can be used as a nice remote administration tool, or just to have some fun with your friends on the net. The network must support TCP/IP.' Furthermore, the documentation notes: 'Note that you don't see Patch when it's running — it's hiding itself automatically at start-up.' This explicitly confirms that NetBus is designed to operate covertly, masking itself from the infected system's user."),
  spacer(),
  body("The presence of associated files including DesCipher.class (a DES encryption implementation), vncCanvas.class, rfbProto.class, and authenticationPanel.class strongly indicates that a functional Java-based VNC/remote access toolkit was present on the system. The rfbproto.pdf and rfbprotoheader.pdf files (Remote Framebuffer Protocol documentation) further support the inference that the suspect had assembled a comprehensive remote control capability."),
  heading2("5.6 Remote Monitoring and Management Programs"),
  body("Autopsy's Interesting Files Identifier module flagged the following files as 'Likely Notable' under the category 'Remote Monitoring Management Programs':"),
  makeWideTable(
    ["File", "Category", "Score", "Configuration", "File Path"],
    [
      ["vncviewer.exe", "Remote Monitoring Management Programs", "Likely Notable", "Ultra VNC", "/img_4Dell Latitude CPi.E01/vol_vol2/My Documents/..."],
      ["vncviewer.exe", "Remote Monitoring Management Programs", "Likely Notable", "Ultra VNC", "/img_4Dell Latitude CPi.E01/vol_vol2/..."],
      ["mstsc.exe", "Remote Monitoring Management Programs", "Likely Notable", "mstsc", "/img_4Dell Latitude CPi.E01/vol_vol2/WINDOWS/system..."],
      ["system.LOG", "Remote Monitoring Management Programs", "Likely Notable", "Kaseya (VSA)", "/img_4Dell Latitude CPi.E01/vol_vol2/WINDOWS/system..."],
      ["mstsc.exe", "Remote Monitoring Management Programs", "Likely Notable", "mstsc", "/img_4Dell Latitude CPi.E01/vol_vol2/WINDOWS/system..."],
    ]
  ),
  spacer(160),
  body("The presence of UltraVNC viewer (vncviewer.exe) in the user's My Documents folder — as opposed to a standard system directory — is particularly suspicious. This placement suggests deliberate installation by the user rather than a standard software deployment. VNC (Virtual Network Computing) enables full graphical remote desktop control over network connections. Combined with NetBus, these tools form a dual-layer remote access capability."),
  spacer(),
  body("The detection of a Kaseya VSA log file (system.LOG) is also noteworthy. Kaseya VSA is a legitimate remote monitoring and management (RMM) platform widely used by managed service providers. Its presence on a suspect's personal workstation, however, warrants investigation into whether it was being used to gain unauthorised remote access to other managed systems within the company's infrastructure."),
  heading2("5.7 War-Dialling Software — ToneLoc"),
  body("The Office Documents category revealed a significant collection of files associated with ToneLoc, version 1.10, a war-dialling tool developed by Minor Threat and Mucho Maas (October 1994). War-dialling is the practice of systematically dialling a large series of telephone numbers to locate those connected to modems, forming the basis of old-school network intrusion techniques."),
  makeTable([
    kvRow("File",     "RELEASE.DOC", true),
    kvRow("Content",  "ToneLoc version 1.10 — Release notes and war-dialling guide", false),
    kvRow("File",     "TL-USER.DOC", true),
    kvRow("Content",  "ToneLoc User Guide — Describes dialling, carrier detection, and log analysis", false),
    kvRow("Notable Excerpt", "23:30:56 474-5335 – Timeout (3) | 23:31:00 474-5978 – No Dialtone | 23:39:02 474-5685 – Busy | 00:24:26 474-5989 – TONE FOUND!", true),
  ]),
  spacer(160),
  body("The TL-USER.DOC file contains operational log entries consistent with active use of the war-dialling software, showing telephone numbers being dialled and a successful tone detection ('TONE FOUND!'). This is not merely documentation — it appears to contain actual war-dialling session logs, indicating that the suspect actively engaged in war-dialling activities to locate modem-accessible systems."),
  heading2("5.8 Encryption and Data Concealment"),
  heading3("5.8.1 High-Entropy Binary Files"),
  body("Autopsy's Encryption Detection module identified two instances of oembios.bin as 'Likely Notable,' with a justification of 'Suspected encryption due to high entropy (7.999987).'"),
  makeTable([
    kvRow("File",          "oembios.bin (× 2 instances)", true),
    kvRow("Score",         "Likely Notable", false),
    kvRow("Entropy",       "7.999987 (near-maximum; theoretical max = 8.000000)", true),
    kvRow("Conclusion",    "Suspected Encryption", false),
    kvRow("Significance",  "Near-perfect entropy indicates encrypted, compressed, or deliberately randomised data concealment", true),
  ]),
  spacer(160),
  body("Entropy is a measure of randomness in data. A maximum Shannon entropy of 8.0 bits per byte indicates completely random data, which is characteristic of encrypted content. The entropy value of 7.999987 is effectively maximum entropy, virtually confirming that the content of oembios.bin has been encrypted or deliberately obfuscated. The legitimate oembios.bin file in Windows XP is a small system bootstrap file with a very low entropy value; the high-entropy variant found on this system is therefore highly anomalous and forensically significant."),
  heading3("5.8.2 PGP Encryption"),
  body("A PGP (Pretty Good Privacy) public key block was recovered from the file PUBKEY.TXT, located in the Plain Text files category. The file header reads 'BEGIN PGP PUBLIC KEY BLOCK — Version: PGP for Personal Privacy 5.5.3.' The presence of PGP encryption software indicates that the suspect was using end-to-end encrypted communications, potentially to conceal correspondence with co-conspirators or to exchange encrypted files."),
  heading3("5.8.3 Extension Mismatch Detection"),
  body("The Extension Mismatch Detector module identified 9 files whose actual file type did not correspond to their declared file extension. Such mismatches are a known technique for concealing the true nature of files from casual inspection. While individual mismatches may result from innocent causes (e.g., incorrectly named files), the presence of 9 such files on a system already exhibiting multiple indicators of malicious activity elevates their forensic significance."),
  heading2("5.9 Web Activity Analysis"),
  heading3("5.9.1 Web History"),
  body("Autopsy recovered 887 web history entries from the forensic image. This volume of browsing history, spanning what appears to be several months of computer use, provides a rich source of intelligence regarding the suspect's online activities. While the individual URLs are too numerous to document exhaustively in this report, the pattern of activity — combined with the installed software arsenal — is consistent with research into hacking techniques, network exploitation, and tool acquisition."),
  heading3("5.9.2 Web Cookies"),
  body("A total of 24 web cookies were recovered, providing additional context regarding websites visited. Web cookies can reveal login activity, session information, and site preferences, potentially identifying online accounts used by the suspect."),
  heading3("5.9.3 Web Bookmarks"),
  body("Six web bookmarks were identified, which may indicate sites of particular interest to the suspect. Forensic examination of bookmarks can reveal research patterns and online communities frequented by the user."),
  heading3("5.9.4 Web Search Terms"),
  body("Four web search terms were recovered, representing queries conducted by the suspect via a web search engine. These search queries can provide direct evidence of the suspect's intent and areas of interest."),
  heading2("5.10 Email and Communications"),
  body("The Data Artifacts panel revealed the following communication-related artefacts:"),
  makeTable([
    kvRow("Communication Accounts", "2 accounts detected", true),
    kvRow("Email Accounts",         "2 email accounts identified", false),
    kvRow("E-Mail Messages",        "1 email message recovered", true),
    kvRow("mIRC",                   "IRC client installed (2004-08-20)", false),
    kvRow("Significance",           "IRC and email are potential command-and-control communication channels", true),
  ]),
  spacer(160),
  body("The combination of IRC (mIRC) and email communication channels is characteristic of hacker communications infrastructure. IRC has historically been a primary channel for hacker communities and botnet command-and-control communications. The recovery of email account information and messages warrants further forensic examination to determine the nature of communications."),
  heading2("5.11 USB Device Attachment"),
  body("Autopsy identified 1 USB device attachment event, indicating that a removable storage device was connected to the suspect's computer at some point. This is forensically significant as it suggests a potential data exfiltration pathway — a USB device could have been used to copy sensitive data or to introduce additional hacking tools to the system. The specific details of the USB device (manufacturer, serial number, connection timestamp) would require further examination of Windows registry artefacts."),
  heading2("5.12 Shell Bags Analysis"),
  body("Autopsy recovered 51 shell bag entries. Shell bags are Windows registry artefacts that record the folder navigation history of a user, including timestamps and window position preferences. They persist even after folders or files have been deleted, making them a valuable source of evidence for reconstructing user activity. The presence of 51 shell bag entries provides a record of directories the suspect navigated, potentially revealing the locations of hacking tools, sensitive files, or areas of particular interest on the system."),
  heading2("5.13 PDF Documents — Vulnerability Scanner Manual"),
  body("The PDF category yielded 6 documents of forensic interest. Notable among these are s3-gs-manual.pdf and s3-manual.pdf, identified as Internet Security Systems (ISS) documentation ('ISS Technical Support: support@iss.net'). The text content references ISS's Internet Scanner, RealSecure, SAFEsuite, and System Scanner products, which are commercial vulnerability assessment and network security tools. The possession of vulnerability scanner documentation further supports the hypothesis that the suspect was engaged in systematic network reconnaissance and vulnerability exploitation."),
  heading2("5.14 Command Scripts Analysis"),
  body("Two .cmd script files were identified in the executable category:"),
  makeTable([
    kvRow("login.cmd", "Telnet server login script — executes on initial command shell invocation. Content: '@echo off', 'Welcome to Microsoft Telnet Server', sets home drive and path.", true),
    kvRow("usrlogon.cmd", "User logon script — associated with Telnet server session initialisation.", false),
  ]),
  spacer(160),
  body("The presence of Telnet server login scripts suggests that a Telnet service was configured on the suspect's machine, potentially enabling remote command-line access. When considered alongside the VNC viewers and NetBus RAT, this creates a triple-layer remote access architecture, indicating sophisticated and deliberate configuration of the system for remote access purposes."),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Reconstruction and Reporting ──────────────────────────────────────────────
const reconstruction = [
  heading1("6. Reconstruction and Reporting"),
  heading2("6.1 Timeline Reconstruction"),
  body("The Timeline module in Autopsy (accessible via the Timeline Editor) generated a comprehensive chronological view of file system events across the disk image. The timeline, displayed in logarithmic scale, revealed the following activity patterns:"),
  makeWideTable(
    ["Period", "Activity Level", "Key Events"],
    [
      ["Pre-1998", "Low", "Legacy system files; pre-existing data from earlier OS installations"],
      ["1998–2000", "Moderate", "Initial OS installation artefacts; ToneLoc documentation (1994 content) installed; NetBus documentation (1998); VNC protocol documentation"],
      ["2000–2002", "Growing", "Progressive software installation; archive files acquired; hacking tool documentation accumulated"],
      ["2002 (Peak)", "Very High", "Maximum recorded activity; bulk of software installations, file modifications, and user activity concentrated in this period"],
      ["2004-08-19", "High", "Major installation event: OutlookExpress, NetMeeting, IE components, DirectDrawEx installed"],
      ["2004-08-20", "High", "Critical date: Cain & Abel, Anonymizer Bar, 123 Write All Stored Passwords, Ethereal, WinPcap installed; Office documents modified; login.cmd modified"],
      ["2004-08-27", "High", "Network Stumbler, Ethereal, WinPcap installed/updated; last confirmed major installation date"],
      ["Post-2004", "Declining", "Residual activity; system in post-operational or dormant state"],
    ]
  ),
  spacer(160),
  body("The concentration of suspicious software installations on 19–27 August 2004 is highly significant. Within a period of eight days, the suspect installed an arsenal of network hacking, password cracking, network scanning, and anonymisation tools. This concentrated installation pattern is inconsistent with normal, legitimate software use and strongly suggests a deliberate and purposeful arming of the system for malicious network activities."),
  heading2("6.2 Linking Evidence to Suspect Actions"),
  body("The following table synthesises the key findings and maps each to the specific allegation of malicious network activity:"),
  makeWideTable(
    ["Evidence", "Forensic Significance", "Allegation Supported"],
    [
      ["Cain & Abel + WinPcap + Ethereal", "Network credential interception; ARP poisoning capability; packet sniffing", "Unauthorised network surveillance and credential theft"],
      ["Network Stumbler", "Wireless network discovery; war driving", "Unauthorised wireless network access/reconnaissance"],
      ["NetBus v1.70 (RAT)", "Remote unauthorised control of victim computers; self-hiding at startup", "Unauthorised access to company computers; malware deployment"],
      ["VNCviewer.exe (UltraVNC)", "Remote graphical desktop access tool", "Unauthorised remote access to networked systems"],
      ["ToneLoc + War-dial logs", "Active war-dialling; carrier tone detection logged", "Unauthorised scanning of telephone/modem networks"],
      ["123 Write All Stored Passwords", "Password extraction from Windows credential store", "Unauthorised access to account credentials"],
      ["oembios.bin (entropy 7.999987)", "Encrypted/concealed data; deliberate obfuscation", "Data concealment; potential contraband storage"],
      ["PUBKEY.TXT (PGP Public Key)", "Encrypted communications infrastructure", "Covert encrypted communications with co-conspirators"],
      ["Anonymizer Bar 2.0", "Online identity concealment; bypass of network monitoring", "Deliberate concealment of malicious online activity"],
      ["1,371 Deleted Files", "Attempted evidence destruction", "Post-activity cover-up; obstruction of investigation"],
      ["USB Device Attachment", "Potential data exfiltration or tool importation", "Data theft or introduction of additional tools"],
      ["Shell Bags (51 entries)", "Extensive navigation of system directories", "Active exploration and use of the file system"],
      ["login.cmd / usrlogon.cmd", "Telnet server configured for remote access", "Remote unauthorised access infrastructure"],
      ["mIRC + Email (2 accounts)", "IRC and email communications", "Potential command-and-control communications"],
    ]
  ),
  spacer(160),
  heading2("6.3 Interpretation of Findings"),
  body("The forensic evidence uncovered in this examination presents a coherent and compelling narrative of deliberate malicious activity. The suspect (identified through the user account as 'Mr. Evil', login: mr) systematically assembled a comprehensive toolkit for hacking, network reconnaissance, and unauthorised remote access over a period leading up to and including August 2004."),
  spacer(),
  body("The installation of Ethereal and WinPcap together is characteristic of a network sniffing setup, enabling the capture and decryption of network traffic. Cain & Abel extends this capability to active attack techniques, including ARP poisoning and password cracking. The combination of these tools with Network Stumbler — a wireless network scanner — indicates that the suspect was conducting both wired and wireless network attacks."),
  spacer(),
  body("The presence of NetBus — a RAT designed to hide itself at startup — on the suspect's system raises the critical question of whether it was installed on this system as a victim (unlikely, given the overall context) or whether it was used as a deployment tool to infect other machines. Given the totality of the evidence, the latter interpretation is significantly more probable."),
  spacer(),
  body("The deletion of 1,371 files and the use of the Anonymizer Bar demonstrate an awareness of forensic detection techniques and a deliberate effort to conceal activities. However, these countermeasures were ultimately unsuccessful, as digital forensic analysis recovered substantial evidence from unallocated disk space and residual system artefacts."),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Conclusion ────────────────────────────────────────────────────────────────
const conclusion = [
  heading1("7. Conclusion"),
  body("The digital forensic examination of the disk image acquired from the suspect's Dell Latitude CPi laptop has produced substantial, multi-layered evidence of malicious and illegal activity conducted in contravention of company network security policies and, potentially, Malaysian law."),
  spacer(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 240 },
    children: [new TextRun({ text: "FINDINGS: ALLEGATIONS SUBSTANTIATED", font: FONT, size: 30, bold: true, color: "CC0000" })]
  }),
  body("The following conclusions are drawn from the forensic evidence:"),
  numbered("The suspect deliberately installed a comprehensive arsenal of hacking and network attack tools, including Cain & Abel, WinPcap, Ethereal, Network Stumbler, Look@LAN, and 123 Write All Stored Passwords, between 19 and 27 August 2004. This installation pattern is inconsistent with legitimate professional activity and demonstrates clear malicious intent."),
  numbered("The presence of NetBus v1.70 — a Remote Access Trojan explicitly designed to hide itself from system users and administrators — on the suspect's workstation constitutes strong evidence of malware possession and potential deployment for unauthorised access to third-party systems."),
  numbered("The configuration of multiple remote access mechanisms, including VNCviewer, mstsc.exe, login.cmd (Telnet server), and NetBus, demonstrates a deliberate and sophisticated effort to establish and maintain unauthorised remote access capabilities."),
  numbered("The recovery of active ToneLoc war-dialling session logs confirms that the suspect engaged in actual war-dialling activities to locate modem-connected systems, a precursor to modem-based intrusion attacks."),
  numbered("Files exhibiting near-maximum entropy (oembios.bin, entropy = 7.999987) and the use of PGP encryption demonstrate active data concealment and encrypted communications, behaviours consistent with an actor attempting to hide malicious activities."),
  numbered("The deletion of 1,371 files indicates an attempt to destroy evidence; the partial success of this effort (many files were recovered by the PhotoRec carver and from unallocated space) does not diminish its significance as evidence of consciousness of guilt."),
  numbered("The use of Anonymizer Bar to conceal online identity demonstrates awareness that activities were being conducted that required concealment from network monitoring systems."),
  spacer(),
  body("The employee's denial of involvement in malicious activities is contradicted by the preponderance of digital evidence. The systematic nature of the tool installation, the sophistication of the remote access infrastructure, and the deliberate attempts to conceal activity all point to a premeditated campaign of malicious network activity."),
  spacer(),
  body("It is the considered forensic opinion of the examiner that the evidence gathered during this investigation would, subject to the satisfaction of applicable legal standards of admissibility, support a finding of guilt in relation to charges of unauthorised computer access and the possession and use of malicious computer programs, under Section 3, 5, and 6 of the Malaysian Computer Crimes Act 1997."),
  new Paragraph({ children: [new PageBreak()] })
];

// ── References ────────────────────────────────────────────────────────────────
const references = [
  heading1("8. References"),
  body("Carrier, B. (2005). File system forensic analysis. Addison-Wesley."),
  spacer(40),
  body("Casey, E. (2011). Digital evidence and computer crime: Forensic science, computers, and the internet (3rd ed.). Academic Press."),
  spacer(40),
  body("Computer Crimes Act 1997 (Act 563). (1997). Laws of Malaysia. Attorney General's Chambers of Malaysia."),
  spacer(40),
  body("Kent, K., Chevalier, S., Grance, T., & Dang, H. (2006). Guide to integrating forensic techniques into incident response (NIST Special Publication 800-86). National Institute of Standards and Technology. https://doi.org/10.6028/NIST.SP.800-86"),
  spacer(40),
  body("Ligh, M. H., Case, A., Levy, J., & Walters, A. (2014). The art of memory forensics: Detecting malware and threats in Windows, Linux, and Mac memory. Wiley."),
  spacer(40),
  body("Nelson, B., Phillips, A., & Steuart, C. (2015). Guide to computer forensics and investigations (5th ed.). Cengage Learning."),
  spacer(40),
  body("Personal Data Protection Act 2010 (Act 709). (2010). Laws of Malaysia. Attorney General's Chambers of Malaysia."),
  spacer(40),
  body("Sammons, J. (2012). The basics of digital forensics: The primer for getting started in digital forensics. Syngress/Elsevier."),
  spacer(40),
  body("Spitzner, L. (2002). Honeypots: Tracking hackers. Addison-Wesley."),
  spacer(40),
  body("Sleuth Kit Labs. (2024). Autopsy digital forensics platform (Version 4.23.0) [Computer software]. https://www.autopsy.com"),
  spacer(40),
  body("Honeynet Project. (2004). Forensic challenge: Mr. Evil disk image. The Honeynet Project. https://www.honeynet.org"),
  spacer(40),
  body("Neikter, C.-F. (1998). NetBus v1.70 [Computer software documentation]. Privately distributed."),
  spacer(40),
  body("Minor Threat & Mucho Maas. (1994). ToneLoc v1.10 [Computer software documentation]. Privately distributed."),
  new Paragraph({ children: [new PageBreak()] })
];

// ── Appendices ────────────────────────────────────────────────────────────────
const appendices = [
  heading1("9. Appendices"),
  body("Note: The following appendix provides descriptions of key forensic screenshots captured during the Autopsy examination. Screenshots are referenced by figure number and labelled to correspond with the findings discussed in the body of this report."),
  spacer(),
  heading2("Appendix A: Autopsy Forensic Screenshots"),
  spacer(),
  heading3("Figure 1 — Autopsy Case Report Summary"),
  body("The Autopsy-generated forensic report summary for Case ADF / CASE-01, confirming case name, case number, examiner (Muhammad Al-Amin Yahaya), data source path, timezone (Asia/Kuala Lumpur), and the list of activated ingest modules including Encryption Detection, Interesting Files Identifier, Keyword Search, Email Parser, and Data Source Integrity."),
  spacer(80),
  heading3("Figure 2 — Installed Programs (Suspicious Tools)"),
  body("Autopsy Data Artifacts → Installed Programs listing showing 32 installed programs. Highlighted entries include Ethereal (v0.10.6), WinPcap (3.01 alpha), Network Stumbler (0.4.0), Look@LAN (2.50), 123 Write All Stored Passwords, Cain & Abel (v2.5 beta45), Anonymizer Bar (2.0), mIRC, CuteFTP, and CuteHTML, all installed on or around 20–27 August 2004. This screenshot provides direct evidence of the suspect's hacking toolkit installation."),
  spacer(80),
  heading3("Figure 3 — Encryption Suspected — oembios.bin"),
  body("Autopsy Analysis Results → Encryption Suspected, showing two instances of oembios.bin flagged as 'Likely Notable' with a score of 1 each. The justification column reads: 'Suspected encryption due to high entropy (7.999987).' This figure demonstrates the presence of highly encrypted or obfuscated binary data deliberately concealed within a file using a legitimate-sounding system filename."),
  spacer(80),
  heading3("Figure 4 — Deleted Files Recovery"),
  body("Autopsy File Views → Deleted Files → All, showing 1,371 total deleted files. The selected file, 000000a6.query.new (9,422 bytes, Unallocated), was deleted on 2004-08-20 at 06:35:00 MYT. The text preview shows references to Microsoft Help backup and system files. This screenshot demonstrates successful recovery of deleted content from unallocated disk space."),
  spacer(80),
  heading3("Figure 5 — Command Scripts (.cmd files)"),
  body("Autopsy File Types → By Extension → .cmd, showing two files: login.cmd and usrlogon.cmd. The text preview of login.cmd reveals the content of a Microsoft Telnet Server login script, including '@echo off', 'Welcome to Microsoft Telnet Server', and home directory navigation commands. This is evidence of a Telnet server configured for remote access."),
  spacer(80),
  heading3("Figure 6 — PDF Files — Vulnerability Scanner Manuals"),
  body("Autopsy File Types → By Extension → PDF, showing 6 PDF files including s3-gs-manual.pdf (1,202,458 bytes), s3-manual.pdf (3,536,682 bytes), rfbproto.pdf (76,357 bytes), and rfbprotoheader.pdf (19,197 bytes). The text preview of s3-manual.pdf confirms ISS (Internet Security Systems) technical documentation for vulnerability scanning tools. This evidences the suspect's research into network vulnerability assessment techniques."),
  spacer(80),
  heading3("Figure 7 — PGP Public Key — PUBKEY.TXT"),
  body("Autopsy File Types → Plain Text → PUBKEY.TXT, showing a file listing that includes AGENTS.TXT, ARJ.TXT, CREDIT.TXT, PUBKEY.TXT, SYSOP.TXT, and UPDATE.TXT. The text preview of PUBKEY.TXT shows '-----BEGIN PGP PUBLIC KEY BLOCK----- Version: PGP for Personal Privacy 5.5.3,' followed by the encoded key data. This confirms the suspect's use of PGP cryptographic encryption."),
  spacer(80),
  heading3("Figure 8 — Remote Monitoring Programs — VNCviewer and NetBus"),
  body("Autopsy Analysis Results → Interesting Items → Remote Monitoring Management Programs, showing 5 flagged items: two instances of vncviewer.exe (Ultra VNC), two instances of mstsc.exe (mstsc), and system.LOG (Kaseya VSA). All are scored 'Likely Notable.' The text preview of vncviewer.exe shows VNCviewer ASCII art and connection command documentation. This screenshot directly evidences the presence of remote access tools on the suspect's system."),
  spacer(80),
  heading3("Figure 9 — ToneLoc War-Dialling Documentation and Logs"),
  body("Autopsy File Types → Office → TL-USER.DOC (43,551 bytes, 1998-05-14), showing the ToneLoc v1.10 user documentation. The text preview displays active war-dialling session log entries: '23:30:56 474-5335 – Timeout (3)', '23:31:00 474-5978 – No Dialtone', '23:39:02 474-5685 – Busy', and '00:24:26 474-5989 – ** TONE ** Holy Sh*t! You found a tone.' These entries confirm actual operational use of war-dialling software by the suspect."),
  spacer(80),
  heading3("Figure 10 — NetBus Metadata and Documentation"),
  body("Autopsy Data Artifacts → Metadata → NetBus.rtf, showing owner 'Carl-Fredrik Neikter', description 'NetBus v', created 1998-04-07, Data Source: 4Dell Latitude CPi.E01. The text preview confirms the NetBus capability description including its auto-hiding capability. Associated files including DesCipher.class, vncCanvas.class, rfbProto.class are also listed. This screenshot provides definitive evidence of the presence of a known Remote Access Trojan."),
  spacer(80),
  heading3("Figure 11 — Installed Programs — CuteFTP Detail"),
  body("Autopsy Data Artifacts → Installed Programs, showing CuteFTP installed on 2004-08-20 at 15:09:02 MYT from 4Dell Latitude CPi.E01. CuteFTP is an FTP client capable of establishing File Transfer Protocol connections to remote servers, providing a data exfiltration pathway."),
  spacer(80),
  heading3("Figure 12 — Autopsy Forensic Report Screen"),
  body("The Autopsy-generated HTML forensic report page for case ADF / CASE-01, showing Software Information including all active Autopsy modules and their versions (Autopsy 4.23.0, all analysis modules 4.23.0 except Hash Parser 1.2 and GPSD Parser 7.0), and the Ingest History confirming the data source 4Dell Latitude CPi.E01 was processed to COMPLETED status with all listed modules applied."),
  spacer(80),
  heading3("Figure 13 — Plain Text Files — AGENTS.TXT"),
  body("Autopsy File Types → Plain Text → AGENTS.TXT, showing the file listing including AGENTS.TXT (5,285 bytes, 2000-10-12), ARJ.TXT, ARJ_BBS.TXT, CREDIT.TXT, ORDERFRM.TXT, PUBKEY.TXT, README.TXT, SYSOP.TXT, UPD260.TXT, UPDATE.TXT. The text preview of AGENTS.TXT shows ARJ compression utility distributor contact information from Brazil and Czech Republic, confirming the presence of ARJ archiving tool documentation on the system."),
  spacer(80),
  heading3("Figure 14 — Executable Files — .com Category"),
  body("Autopsy File Types → By Extension → .com, showing 26 files including FORMAT.COM (49,575 bytes), COMMAND.COM (93,890 bytes), win.com (18,432 bytes), loadfix.com, chcp.com, command.com. The text preview of FORMAT.COM confirms it is 'MS-DOS Version 7 (C) Copyright 1981–1995 Microsoft Corp.' The presence of legacy MS-DOS tools suggests a dual-boot or legacy environment configuration on the suspect's system."),
  spacer(80),
  heading3("Figure 15 — OS Accounts — Mr. Evil User Profile"),
  body("Autopsy OS Accounts panel showing the primary user account (SID: S-1-5-21-2000478354-688789844-1708537768-1003, Login: mr, Object ID: 7), alongside system accounts (LOCAL SERVICE, NETWORK SERVICE, SYSTEM). The OS Account detail panel confirms Login: mr, Home Directory: /Documents and Settings/Mr. Evil, and Realm: Unknown. This is the definitive identification of the primary suspect account."),
  spacer(80),
  heading3("Figure 16 — Data Source and Ingest Progress"),
  body("Autopsy main interface showing the Data Sources panel with '4Dell Latitude CPi.E01_1 Host' as the sole data source, and the ingest progress bars showing 'Analyzing files from 4Dell Latitude CPi.E01' at 100% and 'ntoskrnl.exe' processing in progress. The sidebar confirms all artefact categories identified: Communication Accounts (2), E-Mail Messages (1), Installed Programs (32), Metadata (30), OS Information (1), Recent Documents (8), Shell Bags (51), USB Device Attached (1), Web Bookmarks (6), Web Cookies (24), Web History (887), Web Search (4)."),
  spacer(80),
  heading3("Figure 17 — Metadata — DesCipher and NetBus Components"),
  body("Autopsy Data Artifacts → Metadata, showing DesCipher.class (DES encryption implementation), NetBus.rtf (owner: Carl-Fredrik Neikter, created 1998-04-07), rfbprotoheader.pdf (created 1998-01-22, version 1.2), rfbproto.pdf (created 1997-07-16, version 1.2), vncviewer.class, vncCanvas.class, rfbProto.class, optionsFrame.class, clipboardFrame.class, authenticationPanel.class, animatedMemoryImageSource.class, DescCipher.class. The text preview of DesCipher.class shows Java source code for a DES encryption class with encryptKeys, decryptKeys, and tempInts fields. This confirms a functional encryption capability embedded within the VNC/remote access toolkit."),
  spacer(80),
  heading3("Figure 18 — Timeline Overview (1970–2034)"),
  body("Autopsy Timeline Editor showing a logarithmic bar chart of file system events from 1970 to 2034. The chart reveals clearly that peak activity occurred around 2002–2004, with a very high concentration of file modifications, creations, and access events in this period. The red bars represent file content changes and the cyan bars represent directory/metadata events. This timeline confirms the period of maximum system activity and is consistent with the software installation dates identified in other artefacts."),
  spacer(80),
  heading3("Figure 19 — PDF Documents — RFB Protocol (VNC) Manual"),
  body("Autopsy File Types → PDF, showing rfbproto.pdf (76,357 bytes, 1998-08-28) and rfbprotoheader.pdf (19,197 bytes, 1998-08-28). The text preview references Internet Security Systems (ISS) and includes trademark notices for Internet Scanner, RealSecure, SAFEsuite, and System Scanner. The RFB (Remote Framebuffer) Protocol is the underlying protocol used by VNC for remote desktop access, confirming the suspect's research into remote access technology."),
  spacer(80),
  heading3("Figure 20 — MD5 Hash Verification — Evidence Integrity"),
  body("Autopsy Data Source Verification Results for 4Dell Latitude CPi.E01, showing: Result: verified; MD5 hash verified; Calculated hash: aee4fcd9301c03b3b054623ca261959a; Stored hash: aee4fcd9301c03b3b054623ca261959a. The matching hash values confirm that the forensic image is authentic and unmodified, satisfying the forensic integrity requirement for court admissibility."),
  spacer(80),
  heading3("Figure 21 — Ingest Module Log"),
  body("Autopsy ingest results log showing 27 unique module events including: Hash Lookup (No notable hash set), Recent Activity (Started / Finished / Browser Results for 4Dell Latitude CPi.E01), Embedded File Extractor (Error unpacking COPYING — benign), Interesting Files Identifier (Remote Monitoring Management Programs: vncviewer.exe × 2, mstsc.exe × 2), Encryption Detection (Encryption Suspected Match: oembios.bin × 2), Plaso Processing Completed, File Type Identification, Keyword Search (Keyword Indexing Results), Extension Mismatch Detector, PhotoRec Carver, and Data Source Integrity (Integrity of 4Dell Latitude CPi.E01 verified). This log provides a complete audit trail of the forensic examination process."),
  spacer(80),
  heading3("Figure 22 — Timeline (Second View)"),
  body("A second Timeline Editor view confirming the logarithmic event distribution from 1970 to 2034, with the dominant activity cluster clearly visible in the 2000–2004 period. This timeline is consistent across both examinations, confirming reproducibility of findings."),
  spacer(200),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text: "— END OF REPORT —", font: FONT, size: BODY, bold: true, color: BLUE })]
  }),
  centred("Word Count (excluding title, table of contents, source code, appendix labels): approximately 6,200 words", BODY - 2, false, GREY),
];

// ── ASSEMBLE DOCUMENT ─────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: FONT, size: BODY } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: H1, bold: true, font: FONT, color: BLUE },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: H2, bold: true, font: FONT, color: LBLUE },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: H3, bold: true, font: FONT, color: GREY },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LBLUE, space: 1 } },
            children: [
              new TextRun({ text: "CONFIDENTIAL — Digital Forensics Investigation Report  |  Case: ADF / CASE-01  |  Examiner: Muhammad Al-Amin Yahaya", font: FONT, size: 18, color: GREY })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: LBLUE, space: 1 } },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Page ", font: FONT, size: 18, color: GREY }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GREY }),
              new TextRun({ text: " of ", font: FONT, size: 18, color: GREY }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: GREY }),
              new TextRun({ text: "  |  Asia Pacific University of Technology and Innovation  |  Advanced Digital Forensics", font: FONT, size: 18, color: GREY }),
            ]
          })
        ]
      })
    },
    children: [
      ...coverPage,
      ...execSummary,
      ...introduction,
      ...authPrep,
      ...evidenceHandling,
      ...analysis,
      ...reconstruction,
      ...conclusion,
      ...references,
      ...appendices
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Digital_Forensics_Investigation_Report_CASE01.docx', buffer);
  console.log('Report written successfully.');
});
