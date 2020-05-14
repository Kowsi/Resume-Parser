import re
import docx2txt
import pandas as pd
import spacy
from spacy.matcher import PhraseMatcher
from spacy.tokens import Span

nlp = spacy.load("en_core_web_sm")

section_title = ['CandidateInformation', 'SummaryText', 'ToolsAndTechnologies', 'WorkExperience',
                 'Education', 'Extra-curricular', 'AwardsAndRecognition']


def convert_docx2txt(file_name):
    temp = docx2txt.process(file_name)
    # text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
    return temp


def read_csv():
    # below is the csv where we have all the keywords, you can customize your own
    return pd.read_csv('./section_title.csv')


def get_section_data(file_name):
    txt = convert_docx2txt(file_name)
    doc = nlp(txt)
    matcher = PhraseMatcher(nlp.vocab)

    section_dict = read_csv()
    section_words = {}

    for section in section_title[1:]:
        section_words[section] = [nlp(text) for text in section_dict[section].dropna(axis=0)]
        matcher.add(section, None, *section_words[section])

    d = []
    matches = matcher(doc)
    print(matches)
    if len(matches) > 0:
        d.append((section_title[0], doc[:matches[0][1]-1]))
    for index, section in enumerate(matches):
        match_id, start, end = section
        rule_id = nlp.vocab.strings[match_id]

        if index == len(matches) - 1:
            span = doc[end:]
        else:
            span = doc[end: matches[index + 1][1] - 1]

        if str(span.text) != '':
            d.append((rule_id, span.text))
        print(rule_id, doc[start:end].text)
        # print('{}    -    {}'.format(rule_id, span.string))

    return d

def print_details(file_name):
    data = get_section_data(file_name)
    for n, d in data:
        print('{}    -    {}'.format(n, str(d)))


def print_entity(txt):
    doc = nlp(txt)


get_work_experience('Novant Health, Charlotte, NC Mar 2019 – Aug 2019 Senior Business Systems Analyst Lead (Supply Chain Technology) Served as a Business Systems Analyst Lead on a program implementing a tissue tracking and inventory management system to be used by Materials Management and Surgical Services areas providing both desktop and handheld capabilities Served as liaison between leadership, stakeholders and technology to elicit, document and analyse regulatory, business and system requirements for changes to business and technical processes, policies and information systems using various techniques such as requirements workshops, storyboards, surveys, site visits, business process descriptions, document analysis, use cases, scenarios, and workflow analysis Facilitated workflow sessions and documented both current and future state workflows and performed gap analysis to reduce project costs and offer more robust requirements Provided user access for the Genesis system (both handheld and desktop applications) Defined test scenarios, wrote test scripts and led a team of analysts in the test process Managed the test matrix and reporting of testing activities to senior leadership Provided leadership and mentoring of project documentation for business analysts and other project staff Interviewed business and technical stakeholders for the purpose of defining business needs and priorities, as well as reconciling conflicts in order clarify specifications and expectations Was administrator for the MS Teams site Managed data for Item Master and Supplier Master data for upload into the system North Highland, Charlotte, NC May 2018 – Feb 2019 Senior Business Analyst (Duke Energy Customer Connect – Commerce Team) Served as Senior Business Analyst on a program implementing an advanced customer experience solution and consolidating Duke Energy’s four legacy billing systems and edge solutions into one customer platform to deliver the universal experience, leveraging SAP’s Hybris Customer Engagement & Commerce solution and its Industry Solution for Utilities on HANA. Currently playing a critical role in the program by developing subject matter expertise on the Commerce functions and the related new software capabilities within the SAP Hybris Suite Supporting the development of Hybris Commerce from a functional perspective on behalf of the program’s functional leadership and business unit stakeholders Defining common future state customer journey and business processes across jurisdictions Defining detailed business requirements Creating functional designs of the new Hybris Commerce software components Configuring the software Defining testing scenarios Executing test scripts Leading requirements workshops Novant Health, Charlotte, NC Feb 2011 – Apr 2018 Senior Business Analyst (Cardiology, Radiology, Oncology, SPD, OR, Business Intelligence) Served as liaison among project stakeholders to elicit, document and analyse regulatory, business and system requirements for changes to business and technical processes, policies and information systems using various techniques such as requirements workshops, storyboards, surveys, site visits, business process descriptions, document analysis, use cases, scenarios, and workflow analysis Provided leadership and mentoring of project documentation for business analysts and other project staff Interviewed business and technical stakeholders for the purpose of defining business needs and priorities, as well as reconciling conflicts in order clarify specifications and expectations Facilitated workflow sessions and documented the current and future state workflows and performed gap analysis to reduce project costs and offer more robust requirements Transformed business requirements into useable system specifications used throughout the software development lifecycle by developers to perform application enhancements Guided teams in the development of Ally Financial, Charlotte, NC Jan 2010– Feb 2011 Senior Business Analyst Lead Served as liaison for a diverse team of international business partners to achieve source to target mapping for the Basel Program and developed and managed all documentation used across the interface by programmers, project managers and data profilers Worked with Data Governance to ensure all elements required for the Basel Program for North American Operations Retail Auto, Retail Mortgage and International Operations Retail programs were documented, submitted and approved for entry into the BACE Business Glossary Gathered, analyzed and documented business requirements using Caliber and performed detailed analysis of data to aid in solving data defects Performed audits on applications to ensure they adhered to corporate policy within the Commercial Workstream Wachovia/Wells Fargo, Charlotte, NC Aug 2006 – Jan 2010 Senior Business Analyst - Officer Documented regulatory workflow processes and organized and led JAD sessions for various projects Successfully implemented compliance policies for Sarbanes Oxley government regulated procedures Created training materials for all SOX compliance reviews in order to ensure reviews were completed properly and coordinated those materials for Business Access Coordinators to use in the training of their managers during reviews Gathered, analyzed and documented business, technical and functional requirements that assisted developers in the development of an automated tool to enable managers/delegates to request, transfer or terminate system access Supervised over 8600 managers within the General Bank during the completion of Access Review System (ARS) Reviews for Applications, Platform Reviews, Database Reviews (DBMS) and Management of Sensitive Data (MSD) Reviews acting as Project Manager and Business Access Coordinator (BAC) Addressed multiple internal and external regulatory (Sarbanes-Oxley and Graham Leach-Bliley Act Acts) and compliance challenges associated with system access for the entire GBG Removed significant deficiency around System Access Reviews and related controls in the General Banking Group Professional Consultant, Charlotte, NC Sept 1998 – Aug 2006 Senior Business Analyst Lead/Manager Worked with and Managed other analysts to solve problems and maximize performance Encouraged collaboration between departments Created training materials for the business analyst team to ensure each analyst was following appropriate procedures for each project Gathered, analyzed and documented business, technical and functional requirements that assisted developers Managed other analysts and worked with senior level managers to complete performance reviews and professional development for analysts Created and implemented procedures Prepared reports Provided coaching and guidance to staff')
#print_details('./Sample Data/Christina_Austin_BA.docx')
#print_details('./Sample Data/Deepak Kumar_Business Analyst_PM-NC.docx')
#print_details('./Sample Data/Manish_ARORA_Profile.docx')