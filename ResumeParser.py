import docx2txt
import spacy
import re
from spacy.matcher import Matcher
from spacy.tokens import Span

# PHONE_NO_PATTERN = re.compile(r"(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4}[\s]+)")


PERSON_PATTERN = [{'POS': 'PROPN', 'ENT_TYPE': 'PERSON'},
                  {'POS': 'PROPN', 'ENT_TYPE': 'PERSON', 'OP': '?'},
                  {'POS': 'PROPN', 'ENT_TYPE': 'PERSON'}]

EMAIL_ID_PATTERN = [{"LIKE_EMAIL": True}]

PHONE_PATTERN_1 = [{"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                   {"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                   {"SHAPE": "dddd"}]

PHONE_PATTERN_2 = [{"ORTH": "("}, {"SHAPE": "ddd"}, {"ORTH": ")"},
                   {"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                   {"SHAPE": "dddd", "LENGTH": {"==": 4}}]

ADDRESS_PATTERN = [{'POS': 'NUM'}, {'POS': 'PROPN'},
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
GIT_URL_PATTERN = [{"LIKE_URL": True, 'ORTH': {'REGEX': r'\s*(github.com)\s*'}}]


class MatchEvent:

    @classmethod
    def full_name_event(cls, matcher, doc, i, matches):
        match_id, start, end = matches[i]
        entity = Span(doc, start, end, label="EVENT")
        if i == 2:
            full_name = doc[matches[0][1]:matches[0][2]]
            if str(entity.text).startswith(str(full_name)):
                matches[0] = matches[i]


class ResumeParser:
    nlp = spacy.load("en_core_web_sm")

    CANDIDATE_INFO = [{'id': 'FullName', 'match_on': MatchEvent.full_name_event, 'pattern': PERSON_PATTERN},
                      {'id': 'Email',       'match_on': None, 'pattern': EMAIL_ID_PATTERN},
                      {'id': 'Address',     'match_on': None, 'pattern': ADDRESS_PATTERN},
                      {'id': 'Phone', 'match_on': None, 'pattern': PHONE_PATTERN_1},
                      {'id': 'Phone', 'match_on': None, 'pattern': PHONE_PATTERN_2},
                      {'id': 'GithubURL',   'match_on': None, 'pattern': GIT_URL_PATTERN},
                      {'id': 'LinkedInURL', 'match_on': None, 'pattern': LINKEDIN_URL_PATTERN}
                      ]

    def __init__(self, file_name):
        self.txt = self.convert_docx2txt(file_name)
        self.doc = ResumeParser.nlp(self.txt)
        self.matcher = Matcher(self.nlp.vocab, validate=True)
        # print(self.txt)
        self.full_name = None
        self.email_id = None
        self.phone_no = None

    def convert_docx2txt(self, file_name):
        temp = docx2txt.process(file_name)
        text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
        return ' '.join(text)

    def get_candidate_info(self):
        # To store Candidate information
        candidate_info_details = {}

        # Adding all the patterns to the Matcher to retrieve the corresponding details
        for index, info in enumerate(ResumeParser.CANDIDATE_INFO):
            self.matcher.add(info['id'], info['match_on'], info['pattern'])
            candidate_info_details[info['id']] = 'Null'

        # Extracting the details from document text
        matches = self.matcher(self.doc)

        # Iterating the result set and store the information
        for match_id, start, end in matches:
            rule_id = str(self.nlp.vocab.strings[match_id])
            if candidate_info_details[rule_id] == 'Null':
                span = self.doc[start:end]
                candidate_info_details[rule_id] = str(span.text)

        #for key in candidate_info_details:
        #   print(key + " : " + candidate_info_details[key])

        return {'CandidateInformation': candidate_info_details}



def main():
    rp = ResumeParser('./Sample Data/Christina_Austin_BA.docx')
    print(rp.get_candidate_info())
    rp = ResumeParser('./Sample Data/Deepak Kumar_Business Analyst_PM-NC.docx')
    print(rp.get_candidate_info())
    rp = ResumeParser('./Sample Data/Manish_ARORA_Profile.docx')
    print(rp.get_candidate_info())


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