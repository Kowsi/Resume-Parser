import docx2txt
import spacy
import re
import pandas as pd
from spacy.matcher import PhraseMatcher
from spacy.matcher import Matcher
from spacy.tokens import Span

# PHONE_NO_PATTERN = re.compile(r"(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4}[\s]+)")

class MatchEvent:
    PERSON_PATTERN      = [{'POS': 'PROPN', 'ENT_TYPE': 'PERSON'},
                          {'POS': 'PROPN', 'ENT_TYPE': 'PERSON', 'OP': '?'},
                          {'POS': 'PROPN', 'ENT_TYPE': 'PERSON'}]

    EMAIL_ID_PATTERN    = [{"LIKE_EMAIL": True}]

    PHONE_PATTERN_1     = [{"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                           {"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                           {"SHAPE": "dddd"}]

    PHONE_PATTERN_2     = [{"ORTH": "("}, {"SHAPE": "ddd"}, {"ORTH": ")"},
                           {"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                           {"SHAPE": "dddd", "LENGTH": {"==": 4}}]

    ADDRESS_PATTERN     = [{'POS': 'NUM'}, {'POS': 'PROPN'},
                           {'POS': 'PROPN', 'OP': "*"},
                           {'IS_PUNCT': True, 'OP': '?'},
                           {'LEMMA': {'IN': ['apt', 'unit']}, 'OP': '?'},
                           {'POS': 'NUM'},
                           {'ORTH': ',', 'OP': "?"},
                           {'POS': 'PROPN'},
                           {'ORTH': {
                               'REGEX': '(AK|AL|AR|AZ|CA|CO|CT|DC|DE|FL|GA|GU|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VI|VT|WA|WI|WV|WY)'}}
                           # ,{'IS_DIGIT': True, 'LENGTH': 5}
                           ]

    LINKEDIN_URL_PATTERN = [{"LIKE_URL": True, 'ORTH': {'REGEX': r'\s*(linkedin.com)\s*'}}]

    GIT_URL_PATTERN     = [{"LIKE_URL": True, 'ORTH': {'REGEX': r'\s*(github.com)\s*'}}]

    @classmethod
    def full_name_event(cls, matcher, doc, i, matches):
        match_id, start, end = matches[i]
        entity = Span(doc, start, end, label="EVENT")
        if i == 2:
            full_name = doc[matches[0][1]:matches[0][2]]
            if str(entity.text).startswith(str(full_name)):
                matches[0] = matches[i]

    @classmethod
    def summary_text_event(cls, matcher, doc, i, matches):
        match_id, start, end = matches[i]
        entity = Span(doc, start, end, label="EVENT")



class ResumeParser:

    nlp = spacy.load("en_core_web_sm")

    CANDIDATE_INFO = [{'id': 'FullName',    'match_on': MatchEvent.full_name_event, 'pattern': MatchEvent.PERSON_PATTERN},
                      {'id': 'Email',       'match_on': None,                       'pattern': MatchEvent.EMAIL_ID_PATTERN},
                      {'id': 'Address',     'match_on': None,                       'pattern': MatchEvent.ADDRESS_PATTERN},
                      {'id': 'Phone',       'match_on': None,                       'pattern': MatchEvent.PHONE_PATTERN_1},
                      {'id': 'Phone',       'match_on': None,                       'pattern': MatchEvent.PHONE_PATTERN_2},
                      {'id': 'GithubURL',   'match_on': None,                       'pattern': MatchEvent.GIT_URL_PATTERN},
                      {'id': 'LinkedInURL', 'match_on': None,                       'pattern': MatchEvent.LINKEDIN_URL_PATTERN}]

    SECTION_TITLE = ['CandidateInformation', 'SummaryText', 'ToolsAndTechnologies', 'WorkExperience',
                 'Education', 'Extra-curricular', 'AwardsAndRecognition']

    SECTION_INFO_FILE = './section_title.csv'

    PROFILE_INFORMATION = {}

    def __init__(self, file_name):
        self.txt = self.convert_docx2txt(file_name)
        self.doc = ResumeParser.nlp(self.txt)
        self.matcher = Matcher(self.nlp.vocab, validate=True)
        self.section_data = self.load_data(self.doc)

    # Converting Docx to txt using docx2txt
    def convert_docx2txt(self, file_name):
        temp = docx2txt.process(file_name)
        #text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
        #return ' '.join(temp)
        return temp

    def load_data(self, data):
        section_dict = pd.read_csv(self.SECTION_INFO_FILE)
        section_words = {}
        matcher = PhraseMatcher(self.nlp.vocab)
        for section in self.SECTION_TITLE[1:]:
            section_words[section] = [self.nlp(text) for text in section_dict[section].dropna(axis=0)]
            matcher.add(section, None, *section_words[section])

        section_data = {}
        matches = matcher(data)

        if len(matches) > 0:
            section_data[self.SECTION_TITLE[0]] = str(data[:matches[0][1]-1])

        for index, section in enumerate(matches):
            match_id, start, end = section
            rule_id = self.nlp.vocab.strings[match_id]

            if index == len(matches) - 1:
                span = data[end:]
            else:
                span = data[end: matches[index + 1][1] - 1]

            if str(span.text) != '':
                section_data[rule_id] = str(span.text)
            #print(rule_id, data[start:end].text)
            #print('{}    -    {}'.format(rule_id, span.string))

        return section_data

    def get_candidate_info(self, title):
        # To store Candidate information
        candidate_info_details = {}
        data = self.nlp(self.section_data[title])

        # Adding all the patterns to the Matcher to retrieve the corresponding details
        for index, info in enumerate(ResumeParser.CANDIDATE_INFO):
            self.matcher.add(info['id'], info['match_on'], info['pattern'])
            candidate_info_details[info['id']] = 'Null'

        # Extracting the details from document text
        matches = self.matcher(data)

        # Iterating the result set and store the information
        for match_id, start, end in matches:
            rule_id = str(self.nlp.vocab.strings[match_id])
            if candidate_info_details[rule_id] == 'Null':
                span = self.doc[start:end]
                candidate_info_details[rule_id] = str(span.text)

        for key in candidate_info_details:
           print(key + " : " + candidate_info_details[key])

        return {title: candidate_info_details}

    def get_summary_text(self, title):
        data = self.section_data[title]
        data = None if data is None else re.sub(r"\s+", " ", data).strip()
        return {title : data}

    def get_work_experience(self, title):
        data = self.section_data[title]
        data = None if data is None else list(filter(None, data.split('\n\n\n')))
        for i, d in enumerate(data):
            print('{}----{}'.format(i,re.sub(r"\s+", " ", d).strip()))
        return {title : data}

    def parse_information(self):
        details = self.get_candidate_info(self.SECTION_TITLE[0])
        details.update(self.get_summary_text(self.SECTION_TITLE[1]))
        details.update(self.get_work_experience(self.SECTION_TITLE[3]))
        return details


def main():
    rp = ResumeParser('./Sample Data/Christina_Austin_BA.docx')
    print(rp.parse_information())
    rp = ResumeParser('./Sample Data/Deepak Kumar_Business Analyst_PM-NC.docx')
    print(rp.parse_information())
    rp = ResumeParser('./Sample Data/Manish_ARORA_Profile.docx')
    print(rp.parse_information())


main()

'''
    def get_basicinfo(self):
        for sent in self.doc.sents:
            if self.name is None:
                name = [ee for ee in sent.ents if ee.label_ == 'PERSON']
                self.name = name[0] if name != None else None
                print('Name :' + str(self.name))
                continue

        if self.email_id is None:
            self.email_id = self.get_matcher("EMAIL_ID", EMAIL_ID_PATTERN, self.doc)
            print('Email Id :' + str(self.email_id))

        if self.phone_no is None:
            for match in re.finditer(PHONE_NO_PATTERN, sent.text):
                start, end = match.span()
                self.phone_no = sent.text[start:end]
                print(f" Phone No: '{sent.text[start:end]}'")
        # print(sent)
        
        
    def get_matcher(self, call_back, info, text):
        self.matcher.add(info.id, call_back, info.pattern)
        matches = self.matcher(self.nlp(text))
        # print([t for t in matches])
        # return self.doc[matches[0][1]:matches[0][2]]
        for match_id, start, end in matches:
            span = self.doc[start:end]
            return span.text
        return None

    '''
